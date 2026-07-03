import React, { useEffect, useMemo, useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import "./App.css";

type JwtClaims = {
  name?: string;
  preferred_username?: string;
  upn?: string;
  oid?: string;
  tid?: string;
  aud?: string;
  scp?: string;
  exp?: number;
  iat?: number;
};

export default function App() {
  const [status, setStatus] = useState("Starting...");
  const [token, setToken] = useState("");
  const [claims, setClaims] = useState<JwtClaims | null>(null);
  const [error, setError] = useState("");
  const [insideTeams, setInsideTeams] = useState(false);

  function decodeJwt(jwt: string): JwtClaims | null {
    try {
      const payload = jwt.split(".")[1];
      const decoded = atob(payload.replace(/-/g, "+").replace(/_/g, "/"));

      return JSON.parse(
        decodeURIComponent(
          decoded
            .split("")
            .map((c) => {
              return "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2);
            })
            .join("")
        )
      );
    } catch {
      return null;
    }
  }

  function formatUnixTime(value?: number) {
    if (!value) return "Not available";
    return new Date(value * 1000).toLocaleString();
  }

  async function signInWithTeamsSso() {
    setError("");
    setToken("");
    setClaims(null);
    setStatus("Initializing inside Microsoft Teams...");

    try {
      await microsoftTeams.app.initialize();

      const context = await microsoftTeams.app.getContext();
      setInsideTeams(Boolean(context?.app?.host?.name || context?.user));

      setStatus("Requesting SSO token from Microsoft Teams...");

      const authToken = await microsoftTeams.authentication.getAuthToken();
      const parsedClaims = decodeJwt(authToken);

      setToken(authToken);
      setClaims(parsedClaims);
      setStatus("Success: Teams returned an SSO token for the signed-in user.");
    } catch (e: any) {
      setStatus("SSO did not complete.");
      setError(
        e?.message ||
          "SSO failed. This page must be opened inside Microsoft Teams and the Teams manifest must include webApplicationInfo."
      );
    }
  }

  useEffect(() => {
    signInWithTeamsSso();
  }, []);

  const userName = useMemo(() => {
    return claims?.name || claims?.preferred_username || claims?.upn || "Teams user";
  }, [claims]);

  const shortToken = token ? `${token.slice(0, 32)}...${token.slice(-24)}` : "";

  return (
    <div className="page">
      <main className="container">
        <section className="hero card">
          <div>
            <p className="label">Microsoft Teams SSO Learning Demo</p>
            <h1>How Microsoft Teams Single Sign-On Works</h1>
            <p className="description">
              This demo shows how a Teams tab app can automatically identify the current
              signed-in Microsoft Teams user without asking the user to sign in again.
            </p>
          </div>

          <div className={token ? "badge success" : error ? "badge error" : "badge loading"}>
            {token ? "SSO Success" : error ? "Needs Setup" : "Testing SSO"}
          </div>
        </section>

        <section className="grid">
          <div className="card panel">
            <p className="panelTitle">Current SSO Status</p>
            <p className="panelValue">{status}</p>
            <p className="small">
              Host check: {insideTeams ? "Microsoft Teams detected" : "Not confirmed as Teams"}
            </p>
          </div>

          <div className="card panel userPanel">
            <p className="panelTitle">Detected User</p>
            <p className="userName">{token ? userName : "Waiting for SSO token..."}</p>
            <p className="small">{claims?.preferred_username || claims?.upn || "No user claim yet"}</p>
          </div>
        </section>

        <section className="card">
          <p className="sectionLabel">The simple explanation</p>
          <h2>What is happening?</h2>

          <div className="flow">
            <div className="flowStep">
              <span>1</span>
              <div>
                <h3>User opens the Teams tab</h3>
                <p>
                  The user is already signed in to Microsoft Teams with a Microsoft 365 account.
                </p>
              </div>
            </div>

            <div className="flowStep">
              <span>2</span>
              <div>
                <h3>The tab initializes the Teams SDK</h3>
                <p>
                  The app calls <code>microsoftTeams.app.initialize()</code> so it can talk to the
                  Teams host environment.
                </p>
              </div>
            </div>

            <div className="flowStep">
              <span>3</span>
              <div>
                <h3>The app asks Teams for a user token</h3>
                <p>
                  The app calls <code>microsoftTeams.authentication.getAuthToken()</code>. Teams then
                  requests or returns a cached Microsoft Entra token for the current user.
                </p>
              </div>
            </div>

            <div className="flowStep">
              <span>4</span>
              <div>
                <h3>The app receives user identity claims</h3>
                <p>
                  The returned token contains claims such as user name, tenant ID, object ID,
                  audience, and expiry time.
                </p>
              </div>
            </div>

            <div className="flowStep">
              <span>5</span>
              <div>
                <h3>Backend validation happens in real apps</h3>
                <p>
                  In production, the frontend should send the token to a backend API. The backend
                  should validate the token before trusting the user identity.
                </p>
              </div>
            </div>
          </div>
        </section>

        <section className="twoColumn">
          <div className="card">
            <p className="sectionLabel">Required Teams setup</p>
            <h2>Manifest requirement</h2>
            <p>
              The Teams app manifest must include <code>webApplicationInfo</code>. This connects
              the Teams app to a Microsoft Entra app registration.
            </p>

            <pre className="codeBlock">{`"webApplicationInfo": {
  "id": "<MICROSOFT_ENTRA_APP_CLIENT_ID>",
  "resource": "api://<YOUR_DOMAIN>/<MICROSOFT_ENTRA_APP_CLIENT_ID>"
}`}</pre>
          </div>

          <div className="card">
            <p className="sectionLabel">Required Entra setup</p>
            <h2>Microsoft Entra requirement</h2>
            <p>
              The Microsoft Entra app registration should expose an API scope, usually named:
            </p>

            <pre className="codeBlock">{`access_as_user`}</pre>

            <p>
              The Application ID URI must match the <code>resource</code> value in the Teams
              manifest.
            </p>
          </div>
        </section>

        {error && (
          <section className="message errorBox">
            <p className="messageTitle">Why SSO may not work yet</p>
            <p>{error}</p>

            <ul>
              <li>The page may be opened in a normal browser instead of inside Microsoft Teams.</li>
              <li>The Teams manifest may not include <code>webApplicationInfo</code>.</li>
              <li>The Microsoft Entra app registration may not expose the correct API scope.</li>
              <li>The domain in the manifest may not match the hosted app domain.</li>
            </ul>

            <button onClick={signInWithTeamsSso}>Try SSO again</button>
          </section>
        )}

        {token && (
          <section className="message successBox">
            <p className="messageTitle">SSO token received</p>
            <p>
              Teams returned a token. This means the Teams tab successfully requested identity for
              the current signed-in user.
            </p>
            <p className="token">{shortToken}</p>
            <p className="small">
              Demo note: This page only shows a shortened token. Do not store access tokens in
              browser storage.
            </p>
          </section>
        )}

        <section className="card claimsBox">
          <p className="sectionLabel">Decoded token preview</p>
          <h2>What information did Teams return?</h2>

          {claims ? (
            <div className="claimsGrid">
              <div>
                <p className="claimLabel">Name</p>
                <p>{claims.name || "Not available"}</p>
              </div>

              <div>
                <p className="claimLabel">Username</p>
                <p>{claims.preferred_username || claims.upn || "Not available"}</p>
              </div>

              <div>
                <p className="claimLabel">User Object ID</p>
                <p>{claims.oid || "Not available"}</p>
              </div>

              <div>
                <p className="claimLabel">Tenant ID</p>
                <p>{claims.tid || "Not available"}</p>
              </div>

              <div>
                <p className="claimLabel">Audience</p>
                <p>{claims.aud || "Not available"}</p>
              </div>

              <div>
                <p className="claimLabel">Expires</p>
                <p>{formatUnixTime(claims.exp)}</p>
              </div>
            </div>
          ) : (
            <p>No claims yet. Run this page inside Microsoft Teams to request an SSO token.</p>
          )}
        </section>
      </main>
    </div>
  );
}