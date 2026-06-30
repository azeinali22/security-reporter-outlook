# Clean Outlook Junk Reporter

A minimal Outlook add-in with no `keytar` dependency.

## Run locally

```powershell
cd clean-outlook-junk-reporter
npm install
npm start
```

Open https://localhost:5173/taskpane.html once and accept the browser certificate warning.

## Test in Microsoft 365 Admin Center

Upload `manifest.xml` as a custom Office add-in and assign it only to your test account.

## Important

This is Version 1 test mode. Buttons log the report payload to the browser console. Next step is adding an API or mailbox submission.
