# MailFlowCheck 📤 – Email Helper (PowerShell GUI Tool)

---

## Overview

MailFlowCheck is a Windows PowerShell-based GUI tool designed to troubleshoot SMTP outbound email issues for systems like:

- **Voicemail servers**
- **Mail applications**
- **Custom servers that send email via SMTP on port 25**
- **You Could use it to test Scan to Email,But Only If you have the machine you are running this Network subnet allowed for port 25**

With this tool, you can easily:

- 🔍 **Look up MX records** of a domain
- 🔌 **Check if port 25 is open**
- ✉️ **Send a test email** using PowerShell's built-in SMTP
- 📋 **View logs/output** in a single interface

This is a self-contained GUI, making it simple for IT staff, technicians, or admins to validate configurations without switching between command-line tools.

---

## 🔧 Features

- ✅ **Easy-to-use Windows Forms interface**
- 🔍 **DNS MX Record Lookup**
- 🔌 **Port 25 connectivity test** using `Test-NetConnection`
- ✉️ **Send test email** using `Send-MailMessage`
- 🧠 **Auto-enables Telnet** if missing
- 🖥️ **Displays external IP** (for mail relay allowlisting)

---

## 🔐 Prerequisites

- PowerShell 5.1 or higher
- Administrator privileges (to enable Telnet client if required)
- Outbound network access on port 25
- Execution policy set to allow script execution

---

## ▶️ How to Run

1. **Open PowerShell as Administrator**
2. **Allow script execution (if not already enabled):**

   ```powershell
   Set-ExecutionPolicy RemoteSigned -Scope Process
   ```

   Or, to run the script directly:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\ScanToEmailHelper.ps1
   ```

---

## 📤 Sending a Test Email

1. Enter the **email domain** (e.g., `example.com`)
2. Enter a **From** and **To** email address
3. Click **Run Check** to verify DNS and port 25 availability
4. If port 25 is open, click **Send Test Email**
5. Review the output log for success or troubleshooting hints

---

## 💡 Common Use Case

You’re configuring a scan-to-email function or voicemail-to-email on a device or server. You’re unsure whether SMTP relay (over port 25) is working. Use this script to:

- Validate DNS records
- Confirm firewall rules
- Simulate actual email delivery

---

## 🌐 Helpful Notes

- If using Microsoft 365, ensure your external IP is allowed in your SMTP connector.
- Add your external IP (`$externalIP` shown in logs) to your mail relay allowlist.
- More on Microsoft 365 relay setup: [Microsoft 365 SMTP Relay Guide](https://support.itsolver.net/hc/en-au/articles/12267003536655)

---

## 📌 Disclaimer

Use responsibly in your network environment. This tool performs real SMTP connections and may be subject to mail server security rules.
