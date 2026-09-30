# Expense Tracker

A React + Vite personal expense tracker with Google Sheets synchronization, budget tracking, filters, CSV export, dark mode, quick presets, and a smart text parser.

## Run locally

```bash
npm install
npm run dev
```

Then open the local URL shown by Vite.

## Build for production

```bash
npm run build
npm run preview
```

## Deploy with Vercel

1. Push this project to GitHub.
2. Import the GitHub repository into Vercel.
3. Vercel will detect Vite automatically.
4. Deploy.

No environment variable is required for the frontend unless you later choose to configure the Google Apps Script URL through an environment variable.

## Google Sheets

The app contains the Google Apps Script endpoint code inside `src/App.jsx`. Deploy that Apps Script as a Web App and paste its `/exec` URL into the app's Google Sheets connection field.
