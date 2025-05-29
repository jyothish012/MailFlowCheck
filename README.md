# MailFlowCheck📤 -Email Helper – PowerShell GUI Tool

Overview
This PowerShell-based GUI tool is designed to troubleshoot SMTP outbound email issues for systems like:

Voicemail servers

Multi-function printers (MFPs)

Mail applications

Custom servers that send email via SMTP on port 25

With this tool, users can easily:

Look up MX records of a domain

Check if port 25 is open

Send a test email using PowerShell's built-in SMTP capabilities

View logs/output in a single interface

This is a self-contained GUI, making it simple for IT staff, technicians, or admins to validate configurations without switching between command-line tools.

🔧 Features
✅ Easy-to-use Windows Forms interface

🔍 DNS MX Record Lookup

🔌 Port 25 connectivity test using Test-NetConnection

✉️ Send test email using Send-MailMessage

🧠 Auto-enables Telnet if missing

🖥️ Displays external IP (for mail relay allowlisting)

🔐 Prerequisites
PowerShell 5.1 or higher

Administrator privileges (to enable Telnet client if required)

Outbound network access on port 25

Execution policy set to allow script execution

▶️ How to Run
Step 1: Open PowerShell as Administrator
Step 2: Allow script execution (if not already enabled)
powershell
Copy
Edit
Set-ExecutionPolicy RemoteSigned -Scope Process
If needed:

powershell
Copy
Edit
powershell -ExecutionPolicy Bypass -File .\ScanToEmailHelper.ps1
📤 Sending a Test Email
Enter the email domain (e.g., example.com)

Enter a From and To email address

Click "Run Check" to verify DNS and port 25 availability

If port 25 is open, click "Send Test Email"

Review the output log for success or troubleshooting hints

💡 Common Use Case
You’re configuring a scan-to-email function or voicemail-to-email on a device or server. You’re unsure whether SMTP relay (over port 25) is working. Use this script to:

Validate DNS records

Confirm firewall rules

Simulate actual email delivery

🌐 Helpful Notes
If using Microsoft 365, ensure your external IP is allowed in your SMTP connector.

Add your external IP ($externalIP shown in logs) to your mail relay allowlist.

More on Microsoft 365 relay setup:
https://support.itsolver.net/hc/en-au/articles/12267003536655

📌 Disclaimer
Use responsibly in your network environment. This tool performs real SMTP connections and may be subject to mail server security rules.
