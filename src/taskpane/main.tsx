import React, { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

type EmailInfo = {
  subject: string;
  from: string;
  mode: "outlook" | "browser";
};

function getEmailInfo(): EmailInfo {
  const office = (window as any).Office;

  if (!office?.context?.mailbox?.item) {
    return {
      subject: "Browser test message",
      from: "Not available outside Outlook",
      mode: "browser",
    };
  }

  const item = office.context.mailbox.item;
  const from = item.from || item.sender;

  return {
    subject: item.subject || "(No subject)",
    from: from?.emailAddress || from?.displayName || "Unknown sender",
    mode: "outlook",
  };
}

function App() {
  const [email, setEmail] = useState<EmailInfo>({
    subject: "Loading...",
    from: "Loading...",
    mode: "browser",
  });

  const [status, setStatus] = useState("Ready.");

  useEffect(() => {
    const office = (window as any).Office;

    if (office?.onReady) {
      office.onReady(() => {
        setEmail(getEmailInfo());
      });
    } else {
      setEmail(getEmailInfo());
    }
  }, []);

  function report(type: "Junk" | "Spam") {
    setStatus(`${type} report submitted in test mode.`);
    console.log({
      reportType: type,
      subject: email.subject,
      from: email.from,
      reportedAt: new Date().toISOString(),
    });
  }

  return (
    <main className="container">
      <section className="header">
        <div className="shield">🛡️</div>
        <div>
          <h1>Security Reporter</h1>
          <p>Report junk or spam only. No phishing option.</p>
        </div>
      </section>

      <section className="card">
        <div className="label">Selected Email</div>
        <div><strong>Subject:</strong> {email.subject}</div>
        <div><strong>From:</strong> {email.from}</div>
        <div className="mode">Mode: {email.mode === "outlook" ? "Outlook add-in" : "Browser preview"}</div>
      </section>

      <section className="actions">
        <button className="btn junk" onClick={() => report("Junk")}>Report Junk</button>
        <button className="btn spam" onClick={() => report("Spam")}>Report Spam</button>
      </section>

      <section className="status">{status}</section>
    </main>
  );
}

createRoot(document.getElementById("root")!).render(<App />);