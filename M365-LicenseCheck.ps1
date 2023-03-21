# Load assemblies for WinForms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to connect to Microsoft Online Services
function Connect-MSOnlineServices {
    Import-Module MSOnline
    $Credentials = Get-Credential
    Connect-MsolService -Credential $Credentials

    # Populate dropdown box with AccountSku options
    $AccountSkus = Get-MsolAccountSku
    $licensesCB.Items.Clear()
    foreach ($sku in $AccountSkus) {
        $licensesCB.Items.Add($sku.AccountSkuId)
    }
}

function Search-UsersByLicense {
    $selectedLicense = $licensesCB.SelectedItem
    $results = Get-MsolUser -All | Where-Object {($_.licenses).AccountSkuId -match $selectedLicense}
    
    # Display results in DataGridView
    $dataGridView.Rows.Clear()
    foreach ($user in $results) {
        $dataGridView.Rows.Add($user.DisplayName, $user.UserPrincipalName)
    }
}

function Export-UsersByLicense {
    $selectedLicense = $licensesCB.SelectedItem
    $results = Get-MsolUser -All | Where-Object {($_.licenses).AccountSkuId -match $selectedLicense}
    
    # Export results to CSV
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $results | Select-Object DisplayName, UserPrincipalName | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
    }
}

function Clear-DataGridView {
    $dataGridView.Rows.Clear()
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.Text = "MS Online User License Search"

# Connect button
$connectBTN = New-Object System.Windows.Forms.Button
$connectBTN.Location = New-Object System.Drawing.Point(10, 10)
$connectBTN.Size = New-Object System.Drawing.Size(100, 30)
$connectBTN.Text = "Connect"
$connectBTN.Add_Click({Connect-MSOnlineServices})
$form.Controls.Add($connectBTN)

# Dropdown box
$licensesCB = New-Object System.Windows.Forms.ComboBox
$licensesCB.Location = New-Object System.Drawing.Point(120, 10)
$licensesCB.Size = New-Object System.Drawing.Size(400, 30)
$form.Controls.Add($licensesCB)

# Search button
$searchBTN = New-Object System.Windows.Forms.Button
$searchBTN.Location = New-Object System.Drawing.Point(530, 10)
$searchBTN.Size = New-Object System.Drawing.Size(100, 30)
$searchBTN.Text = "Search"
$searchBTN.Add_Click({Search-UsersByLicense})
$form.Controls.Add($searchBTN)

# Export button
$exportBTN = New-Object System.Windows.Forms.Button
$exportBTN.Location = New-Object System.Drawing.Point(640, 10)
$exportBTN.Size = New-Object System.Drawing.Size(100, 30)
$exportBTN.Text = "Export"
$exportBTN.Add_Click({Export-UsersByLicense})
$form.Controls.Add($exportBTN)

# DataGridView
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 50)
$dataGridView.Size = New-Object System.Drawing.Size(760, 500)
$dataGridView.ColumnCount = 2
$dataGridView.Columns[0].Name = "Display Name"
$dataGridView.Columns[1].Name = "User Principal Name"
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$form.Controls.Add($dataGridView)

# Clear button
$clearBTN = New-Object System.Windows.Forms.Button
$clearBTN.Location = New-Object System.Drawing.Point(740, 10)
$clearBTN.Size = New-Object System.Drawing.Size(30, 30)
$clearBTN.Text = "X"
$clearBTN.Add_Click({Clear-DataGridView})
$form.Controls.Add($clearBTN)

# Show form
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
