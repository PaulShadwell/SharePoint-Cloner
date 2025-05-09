"use client";

import "./globals.css";
import { ReactNode } from "react";
import { MsalProvider } from "@azure/msal-react";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig } from "../lib/authConfig";

const msalInstance = new PublicClientApplication(msalConfig);

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="en">
      <body>
        <MsalProvider instance={msalInstance}>{children}</MsalProvider>
      </body>
    </html>
  );
}

