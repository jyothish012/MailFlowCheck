Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Scan-to-Email Helper"
$form.Size = New-Object System.Drawing.Size(500, 400)
$form.StartPosition = "CenterScreen"
$form.BackColor = "White"

# Fonts
$fontLabel = New-Object System.Drawing.Font("Segoe UI", 10)
$fontButton = New-Object System.Drawing.Font("Segoe UI", 9)

# Printer IP Label & Textbox
$lblIP = New-Object System.Windows.Forms.Label
$lblIP.Text = "Printer IP:"
$lblIP.Location = '20,30'
$lblIP.Size = '100,25'
$lblIP.Font = $fontLabel
$form.Controls.Add($lblIP)

$txtIP = New-Object System.Windows.Forms.TextBox
$txtIP.Location = '130,30'
$txtIP.Size = '300,25'
$form.Controls.Add($txtIP)

# Domain Label & Textbox
$lblDomain = New-Object System.Windows.Forms.Label
$lblDomain.Text = "Email Domain:"
$lblDomain.Location = '20,70'
$lblDomain.Size = '100,25'
$lblDomain.Font = $fontLabel
$form.Controls.Add($lblDomain)

$txtDomain = New-Object System.Windows.Forms.TextBox
$txtDomain.Location = '130,70'
$txtDomain.Size = '300,25'
$form.Controls.Add($txtDomain)

# Output Box
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.Location = '20,120'
$outputBox.Size = '440,180'
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($outputBox)

# Run Button
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run Check"
$btnRun.Location = '180,320'
$btnRun.Size = '120,30'
$btnRun.Font = $fontButton
$form.Controls.Add($btnRun)

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

# Function: Run the main check
$btnRun.Add_Click({
        $outputBox.Clear()
        $ip = $txtIP.Text.Trim()
        $domain = $txtDomain.Text.Trim()

        if (-not $ip -or -not $domain) {
            $outputBox.AppendText("ERROR: Please enter both Printer IP and Email Domain.`n")
            return
        }

        Ensure-Telnet

        try {
            $outputBox.AppendText("Looking up MX records for $domain...`n")
            $mxRecords = (Resolve-DnsName -Type MX $domain -ErrorAction Stop | Sort-Object -Property Preference).NameExchange
            $mx = $mxRecords[0]
            $outputBox.AppendText("Using MX: $mx`n")

            $outputBox.AppendText("Testing port 25 to $mx...`n")
            $result = Test-NetConnection -ComputerName $mx -Port 25

            if ($result.TcpTestSucceeded) {
                $outputBox.AppendText("SUCCESS: Port 25 is OPEN on $mx`n")
            }
            else {
                $outputBox.AppendText("FAILURE: Port 25 is BLOCKED to $mx`n")
                $outputBox.AppendText("You may need to allow port 25 for printer IP $ip in the firewall.`n")
            }
        }
        catch {
            $outputBox.AppendText("ERROR: $_`n")
        }
    })

# Run form
[void]$form.ShowDialog()
