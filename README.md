# DG-Google-App-Scripts

All Google Sheet App Script Automations deployed

## Google clasp - The Apps Script CLI 
# Backup Google Sheets App Script Using clasp (Command Line Interface)

Follow these steps to back up your Google Sheets App Script files using the `clasp` tool:

---

## Prerequisites

- Install [Node.js and npm](https://nodejs.org/) if not already installed.

---

## Step 1: Install clasp

Open your terminal or command prompt and run:

```bash
npm install -g @google/clasp
```

---

## Step 2: Login to clasp

Authenticate clasp with your Google account:

```bash
clasp login
```

A browser window will open. Follow the prompts to authorize clasp.

---

## Step 3: Clone an Existing Apps Script Project

To back up an existing Apps Script project, you need its Script ID:

1. Open your Apps Script project in the browser.
2. Click on **Project Settings** (gear icon).
3. Copy the **Script ID**.

Clone the project to your local machine:

```bash
clasp clone <SCRIPT_ID>
```

Replace `<SCRIPT_ID>` with your actual Script ID.

---

## Step 4: Pull the Latest Code

If you already have a local clasp project, pull the latest code:

```bash
clasp pull
```

This will update your local files with the latest code from Google Apps Script.

---

## Step 5: Backup Your Files

Your Apps Script files are now available locally. You can back them up using your preferred method (e.g., git, cloud storage, manual copy).

---

## Step 6: (Optional) Push Local Changes Back

If you make changes locally and want to update the Apps Script project:

```bash
clasp push
```

---

## Additional Commands

- **List Projects:**  
  ```bash
  clasp list
  ```
- **Open Project in Browser:**  
  ```bash
  clasp open
  ```

---

## References

- [clasp GitHub Repository](https://github.com/google/clasp)