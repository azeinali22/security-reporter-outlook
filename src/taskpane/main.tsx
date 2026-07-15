import React, { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

type EmailInfo = {
  subject: string;
  fromName: string;
  fromEmail: string;
  mode: "outlook" | "browser";
};

function getEmailInfo(): EmailInfo {
  const office = (window as any).Office;

  if (!office?.context?.mailbox?.item) {
    return {
      subject: "Not available outside Outlook",
      fromName: "Unknown sender",
      fromEmail: "Not available",
      mode: "browser",
    };
  }

  const item = office.context.mailbox.item;
  const from = item.from || item.sender;

  return {
    subject: item.subject || "(No subject)",
    fromName: from?.displayName || "Unknown sender",
    fromEmail: from?.emailAddress || "Unknown email",
    mode: "outlook",
  };
}

function App() {
  const [email, setEmail] = useState<EmailInfo>({
    subject: "Loading...",
    fromName: "Loading...",
    fromEmail: "Loading...",
    mode: "browser",
  });

  const [reported, setReported] = useState(false);

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

  function reportSpam() {
    const office = (window as any).Office;
    const item = office?.context?.mailbox?.item;

    console.log({
      reportType: "Junk/Spam",
      itemId: item?.itemId,
      subject: email.subject,
      fromName: email.fromName,
      fromEmail: email.fromEmail,
      reportedAt: new Date().toISOString(),
    });

    setReported(true);
  }

  function closePanel() {
    const office = (window as any).Office;

    if (office?.context?.ui?.closeContainer) {
      office.context.ui.closeContainer();
    }
  }

  return (
    <main className="container">
      <header className="header">
        <img
          src="/assets/icon-128.png"
          alt="Report Junk or Spam"
          className="logo"
        />

        <h1>
          Report
          <br />
          Junk/Spam
        </h1>
      </header>

      {!reported ? (
        <>
          <p className="intro">
            Do you wish to report the following email as junk/spam?
          </p>

          <section className="email-card">
            <div className="card-title">SELECTED EMAIL</div>

            <div className="email-field">
              <span className="field-label">From:</span>

              <strong>{email.fromName}</strong>

              <span className="email-address">
                &lt;{email.fromEmail}&gt;
              </span>
            </div>

            <div className="email-field">
              <span className="field-label">Subject:</span>

              <span>{email.subject}</span>
            </div>
          </section>

          <button className="primary-button" onClick={reportSpam}>
            Report as Junk/Spam
          </button>
        </>
      ) : (
        <section className="confirmation">
          <p>Thanks for reporting!</p>

          <p>
            Reporting junk/spam helps improve junk/spam email detection for you
            and others in the future.
          </p>

          <p>
            Any future emails that are sent from
          </p>

          <p className="sender-summary">
            <strong>{email.fromName}</strong>
            <br />
            &lt;{email.fromEmail}&gt;
          </p>

          <p>
            will be automatically blocked and moved to your{" "}
            <strong>Junk Email</strong> folder.
          </p>

          <button className="close-button" onClick={closePanel}>
            Close
          </button>
        </section>
      )}
    </main>
  );
}

createRoot(document.getElementById("root")!).render(<App />);