Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Scan-to-Email Helper"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"
$form.BackColor = "White"

#Find external IP
$externalIP = (nslookup myip.opendns.com resolver1.opendns.com | Select-String "Address:").Line[-1].Split()[-1]
# Fonts (larger and bold for better visibility)
$fontLabel = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$fontButton = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$fontInput = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)

# Domain Label & Textbox
$lblDomain = New-Object System.Windows.Forms.Label
$lblDomain.Text = "Email Domain:"
$lblDomain.Location = New-Object System.Drawing.Point(30, 30)
$lblDomain.Size = New-Object System.Drawing.Size(140, 35)
$lblDomain.Font = $fontLabel
$lblDomain.ForeColor = [System.Drawing.Color]::Black
$form.Controls.Add($lblDomain)

$txtDomain = New-Object System.Windows.Forms.TextBox
$txtDomain.Location = New-Object System.Drawing.Point(180, 30)
$txtDomain.Size = New-Object System.Drawing.Size(370, 35)
$txtDomain.Font = $fontInput
$form.Controls.Add($txtDomain)

# Sender Email Label & Textbox (always visible)
$lblFrom = New-Object System.Windows.Forms.Label
$lblFrom.Text = "Sender Email:"
$lblFrom.Location = New-Object System.Drawing.Point(30, 80)
$lblFrom.Size = New-Object System.Drawing.Size(140, 35)
$lblFrom.Font = $fontLabel
$lblFrom.ForeColor = [System.Drawing.Color]::Black
$form.Controls.Add($lblFrom)

$txtFrom = New-Object System.Windows.Forms.TextBox
$txtFrom.Location = New-Object System.Drawing.Point(180, 80)
$txtFrom.Size = New-Object System.Drawing.Size(370, 35)
$txtFrom.Font = $fontInput
$form.Controls.Add($txtFrom)

# Recipient Email Label & Textbox (always visible)
$lblTo = New-Object System.Windows.Forms.Label
$lblTo.Text = "Recipient Email:"
$lblTo.Location = New-Object System.Drawing.Point(30, 130)
$lblTo.Size = New-Object System.Drawing.Size(140, 35)
$lblTo.Font = $fontLabel
$lblTo.ForeColor = [System.Drawing.Color]::Black
$form.Controls.Add($lblTo)

$txtTo = New-Object System.Windows.Forms.TextBox
$txtTo.Location = New-Object System.Drawing.Point(180, 130)
$txtTo.Size = New-Object System.Drawing.Size(370, 35)
$txtTo.Font = $fontInput
$form.Controls.Add($txtTo)

# Output Box
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.Location = New-Object System.Drawing.Point(30, 180)
$outputBox.Size = New-Object System.Drawing.Size(520, 180)
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($outputBox)

# Run Button
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run Check"
$btnRun.Location = New-Object System.Drawing.Point(150, 380)
$btnRun.Size = New-Object System.Drawing.Size(150, 45)
$btnRun.Font = $fontButton
$form.Controls.Add($btnRun)

# Add a second button for sending a test email
$btnTestEmail = New-Object System.Windows.Forms.Button
$btnTestEmail.Text = "Send Test Email"
$btnTestEmail.Location = New-Object System.Drawing.Point(320, 380)
$btnTestEmail.Size = New-Object System.Drawing.Size(180, 45)
$btnTestEmail.Font = $fontButton
$btnTestEmail.Enabled = $false
$form.Controls.Add($btnTestEmail)

# Function: Check if Telnet is installed
function Ensure-Telnet {
    $telnetInstalled = Get-WindowsOptionalFeature -Online -FeatureName TelnetClient | Where-Object { $_.State -eq 'Enabled' }
    if (-not $telnetInstalled) {
        $outputBox.AppendText("Installing Telnet Client..." + [Environment]::NewLine)
        Enable-WindowsOptionalFeature -Online -FeatureName TelnetClient -All -NoRestart | Out-Null
        Start-Sleep -Seconds 2
        $outputBox.AppendText("Telnet installed." + [Environment]::NewLine)
    }
    else {
        $outputBox.AppendText("Telnet already installed." + [Environment]::NewLine)
    }
}

# Store the last successful MX
$script:lastMX = $null

# Update the Run Check button click event to enable the test email button if port 25 is open
$btnRun.Add_Click({
        $outputBox.Clear()
        $domain = $txtDomain.Text.Trim()
        $from = $txtFrom.Text.Trim()
        $to = $txtTo.Text.Trim()

        if (-not $domain) {
            $outputBox.AppendText("ERROR: Please enter the Email Domain.`n")
            $btnTestEmail.Enabled = $false
            return
        }

        Ensure-Telnet

        try {
            $outputBox.AppendText("Looking up MX records for $domain...`n")
            $mxRecords = (Resolve-DnsName -Type MX $domain -ErrorAction Stop | Sort-Object -Property Preference).NameExchange
            $mx = $mxRecords[0].ToString()
            $outputBox.AppendText("Using MX: $mx`n")

            $outputBox.AppendText("Testing port 25 to $mx...`n")
            $result = Test-NetConnection -ComputerName $mx -Port 25

            if ($result.TcpTestSucceeded) {
                $outputBox.AppendText("SUCCESS: Port 25 is OPEN on $mx from your Device`n")
                $btnTestEmail.Enabled = $true
                $script:lastMX = $mx
                $outputBox.AppendText("You can now test sending an email from this computer using the fields above and the Send Test Email button.`n")
            }
            else {
                $outputBox.AppendText("FAILURE: Port 25 is BLOCKED to $mx`n")
                $outputBox.AppendText("You need to allow port 25 in the firewall.`n")
                $btnTestEmail.Enabled = $false
                $script:lastMX = $null
            }
        }
        catch {
            $outputBox.AppendText("ERROR: $_`n")
            $btnTestEmail.Enabled = $false
            $script:lastMX = $null
        }
    })

# Add click event for the Send Test Email button
$btnTestEmail.Add_Click({
        if (-not $script:lastMX) {
            $outputBox.AppendText("ERROR: No valid MX record found. Please run the check first and ensure port 25 is open.`n")
            return
        }
        $from = $txtFrom.Text.Trim()
        $to = $txtTo.Text.Trim()
        if (-not $from -or -not $to) {
            $outputBox.AppendText("CANCELLED: No sender or recipient email provided.`n")
            return
        }
        $outputBox.AppendText("Sending test email from $from to $to using $script:lastMX...`n")
        try {
            Send-MailMessage -From $from -To $to -Subject "Test Email from Scan-to-Email Helper" `
                -Body "This is a test email sent via $script:lastMX using PowerShell." `
                -SmtpServer $script:lastMX -Port 25 -ErrorAction Stop
            $outputBox.AppendText("SUCCESS: Test email sent.`n")
        }
        catch {
            $outputBox.AppendText("FAILED: Could not send email.`n")
            $outputBox.AppendText("Possible issue with mail relay or Email restriction restrictions.`n")
            $outputBox.AppendText("See this article if you are using M365:`nhttps://support.itsolver.net/hc/en-au/articles/12267003536655-Configure-Scan-to-Email-using-IP-address-based-connector-for-SMTP-relay`n")
            $outputBox.AppendText("Tip: Add your external IP ($externalIP) to your mail relay's allowed list.' in CMD to find your external IP.`n")
        }
    })

# Run form
[void]$form.ShowDialog()