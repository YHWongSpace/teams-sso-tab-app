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

  async function signInWithTeamsSso() {
    setError("");
    setStatus("Initializing inside Microsoft Teams...");

    try {
      await microsoftTeams.app.initialize();

      try {
        const context = await microsoftTeams.app.getContext();
        setInsideTeams(Boolean(context?.app?.host?.name || context?.user));
      } catch {
        setInsideTeams(true);
      }

      setStatus("Requesting SSO token for the signed-in Teams user...");

      const authToken = await microsoftTeams.authentication.getAuthToken();
      const parsedClaims = decodeJwt(authToken);

      setToken(authToken);
      setClaims(parsedClaims);
      setStatus("Signed in automatically with Microsoft Teams SSO.");
    } catch (e: any) {
      setStatus("Not signed in");
      setError(
        e?.message ||
          "SSO failed. This page must be opened as a Microsoft Teams tab and configured with webApplicationInfo in the Teams manifest."
      );
    }
  }

  useEffect(() => {
    signInWithTeamsSso();
  }, []);

  const userName = useMemo(() => {
    return claims?.name || claims?.preferred_username || claims?.upn || "Teams user";
  }, [claims]);

  const shortToken = token ? `${token.slice(0, 28)}...${token.slice(-18)}` : "";

  return (
    <div className="page">
      <main className="card">
        <div className="header">
          <div>
            <p className="label">Microsoft Teams SSO Demo</p>
            <h1>Auto sign-in webpage</h1>
            <p className="description">
              When this page is installed as a Teams tab, it tries to sign in the current Teams user automatically.
            </p>
          </div>

          <div className={token ? "badge success" : error ? "badge error" : "badge loading"}>
            {token ? "Signed in" : error ? "Action needed" : "Loading"}
          </div>
        </div>

        <div className="grid">
          <div className="panel">
            <p className="panelTitle">Status</p>
            <p className="panelValue">{status}</p>
            <p className="small">
              Host: {insideTeams ? "Microsoft Teams detected" : "Browser / not confirmed"}
            </p>
          </div>

          <div className="panel userPanel">
            <p className="panelTitle">User</p>
            <p className="userName">{token ? userName : "Waiting for SSO..."}</p>
            <p className="small">{claims?.preferred_username || claims?.upn || ""}</p>
          </div>
        </div>

        {error && (
          <div className="message errorBox">
            <p className="messageTitle">SSO could not complete</p>
            <p>{error}</p>
            <button onClick={signInWithTeamsSso}>Try again</button>
          </div>
        )}

        {token && (
          <div className="message successBox">
            <p className="messageTitle">Token received from Teams</p>
            <p className="token">{shortToken}</p>
            <p className="small">
              Do not store this token in browser storage. Send it to your backend only when needed and validate it server-side.
            </p>
          </div>
        )}

        <div className="claimsBox">
          <p className="messageTitle">Decoded claims preview</p>
          <pre>
            {claims
              ? JSON.stringify(
                  {
                    name: claims.name,
                    preferred_username: claims.preferred_username,
                    oid: claims.oid,
                    tid: claims.tid,
                    aud: claims.aud,
                    scp: claims.scp,
                    exp: claims.exp,
                  },
                  null,
                  2
                )
              : "No token yet."}
          </pre>
        </div>
      </main>
    </div>
  );
}