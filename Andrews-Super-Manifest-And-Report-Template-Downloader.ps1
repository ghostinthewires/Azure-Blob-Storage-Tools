<#
.SYNOPSIS
  Andrews-Super-Content-Downloader GUI for Manifest & Report Template Download
 
.DESCRIPTION
  A simple GUI wrapper for AZCopy.exe to simplify the process of Downloading Manifest & Report Templates to Local Storage
 
.INPUTS
  None
 
.OUTPUTS
  None
 
.NOTES
  Version:        1.0
  Creation Date:  06/08/2018

.EXAMPLE
  .\Andrews-Super-Manifest-And-Report-Template-Downloader.ps1
#>
$ErrorActionPreference = "Stop"
# GUI Form
function Call-GUI {

[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null

$form = New-Object System.Windows.Forms.Form
$CancelButton = New-Object System.Windows.Forms.Button
$OkButton = New-Object System.Windows.Forms.Button
$groupBox1 = New-Object System.Windows.Forms.GroupBox
$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkBox3 = New-Object System.Windows.Forms.CheckBox
$label5 = New-Object System.Windows.Forms.Label
$button4 = New-Object System.Windows.Forms.Button
$textBox4 = New-Object System.Windows.Forms.TextBox
$label3 = New-Object System.Windows.Forms.Label
$label4 = New-Object System.Windows.Forms.Label
$textBox3 = New-Object System.Windows.Forms.TextBox
$label2 = New-Object System.Windows.Forms.Label
$textBox2 = New-Object System.Windows.Forms.TextBox
$BrowseButton = New-Object System.Windows.Forms.Button
$textBox1 = New-Object System.Windows.Forms.TextBox
$label1 = New-Object System.Windows.Forms.Label
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
$menu = New-Object 'System.Windows.Forms.MenuStrip'
$fileToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
$exitToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
$editToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
$copyToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
$pasteToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
$helpToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
$aboutToolStripMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
$versionMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
$helpMenuItem = New-Object 'System.Windows.Forms.ToolStripMenuItem'
$form_FormClosing = [System.Windows.Forms.FormClosingEventHandler] {$Script:EndScript++}
	
$Form_StateCorrection_Load = {$form.WindowState = $InitialFormWindowState}
	
$Form_Cleanup_FormClosed = {
	try {
		$form.remove_FormClosing($form_FormClosing)
		$form.remove_Load($FormEvent_Load)
		$form.remove_Load($Form_StateCorrection_Load)
		$form.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch [Exception]
		{ }
	}
	
$Script:OptArray = @()

$End_Script = {
 	$Form.Close()
	$Script:EndScript = 1
	}

$BrowseButton_OnClick = {
	$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ShowNewFolderButton = $false}
	[void]$FolderBrowser.ShowDialog()
	$ManifestPath = $FolderBrowser.SelectedPath
	$textBox1.Text = $ManifestPath
	}
	
$button4_OnClick = {
	$FolderBrowser2 = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ShowNewFolderButton = $false}
	[void]$FolderBrowser2.ShowDialog()
	$LogPath = $FolderBrowser2.SelectedPath
	$textBox4.Text = $LogPath
	}

$OkButton_OnClick = {
	$Script:Source = $DropDownBox1.SelectedItem.Value
	$Script:Target = $DropDownBox.SelectedItem.Value
	If ($checkBox1.Checked) {$Script:OptArray += "/S"}
	If ($checkBox3.Checked) {$Script:OptArray += "/Pattern:*.json"}
	If ($checkBox2.Checked) {
	$LogPath = $textBox4.Text
		If (!$LogPath) {$Script:OptArray += "/LOG"}
		Else {$Script:OptArray += "/LOG:`"$LogPath\AzCopyVerbose.log`""}
	}
	$Script:EndScript = 2
	$Form.Close()
}

$OnLoadForm_StateCorrection = {
	$form.WindowState = $InitialFormWindowState
	}

$form.BackColor = 'White'
$form.Size = '616, 335'
$form.Name = "form"
$form.StartPosition = 'CenterScreen'
$form.Text = "Andrews-Super-Manifest-And-Report-Template-Downloader.ps1"
$form.add_FormClosing($form_FormClosing)
$form.add_Load($FormEvent_Load)
$form.KeyPreview = $True
$form.Add_KeyDown({if ($_.KeyCode -eq "Enter"){
	$Script:Source = $DropDownBox1.SelectedItem.Value
	$Script:Target = $DropDownBox.SelectedItem.Value
	$Script:EndScript = 2
	$Form.Close()
	}})
$form.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$form.Close()
	}})
$form.MainMenuStrip = $menu
$form.Controls.Add($menu)

[void]$menu.Items.Add($fileToolStripMenuItem)
[void]$menu.Items.Add($editToolStripMenuItem)
[void]$menu.Items.Add($helpToolStripMenuItem)
$menu.Location = '0, 0'
$menu.Name = "Menu"
$menu.Size = '344, 24'
$menu.TabIndex = 0
$menu.Text = "Menu"
$menu.BackColor = "White"

[void]$fileToolStripMenuItem.DropDownItems.Add($exitToolStripMenuItem)
$fileToolStripMenuItem.Name = "fileToolStripMenuItem"
$fileToolStripMenuItem.Size = '37, 20'
$fileToolStripMenuItem.Text = "File"
$fileToolStripMenuItem.BackColor = "White"

$exitToolStripMenuItem.Name = "exitToolStripMenuItem"
$exitToolStripMenuItem.Size = '152, 22'
$exitToolStripMenuItem.Text = "Exit"
$exitToolStripMenuItem.BackColor = "White"
$exitToolStripMenuItem.add_MouseEnter($menu_MouseEnter)
$exitToolStripMenuItem.add_MouseLeave($menu_MouseLeave)
$exitToolStripMenuItem.add_Click($End_Script)

[void]$editToolStripMenuItem.DropDownItems.Add($copyToolStripMenuItem)
[void]$editToolStripMenuItem.DropDownItems.Add($pasteToolStripMenuItem)
$editToolStripMenuItem.Name = "editToolStripMenuItem"
$editToolStripMenuItem.Size = '39, 20'
$editToolStripMenuItem.Text = "Edit"

$copyToolStripMenuItem.Name = "copyToolStripMenuItem"
$copyToolStripMenuItem.ShortcutKeys = 'Ctrl+C'
$copyToolStripMenuItem.Size = '152, 22'
$copyToolStripMenuItem.Text = "Copy"
$copyToolStripMenuItem.add_Click({
	$ActiveTextbox = $form.ActiveControl
	$ActiveTextbox.Copy()
	})

$pasteToolStripMenuItem.Name = "pasteToolStripMenuItem"
$pasteToolStripMenuItem.ShortcutKeys = 'Ctrl+V'
$pasteToolStripMenuItem.Size = '152, 22'
$pasteToolStripMenuItem.Text = "Paste"
$pasteToolStripMenuItem.add_Click({
	$ActiveTextbox = $form.ActiveControl
	$ActiveTextbox.Paste()
	})

$CancelButton.Location = '332, 260'
$CancelButton.Name = "CancelButton"
$CancelButton.Size = '75, 25'
$CancelButton.TabIndex = 7
$CancelButton.Text = "Cancel"
$CancelButton.FlatStyle = 'System'
$CancelButton.ForeColor = 'Black'
$CancelButton.add_Click($End_Script)

$form.Controls.Add($CancelButton)

$OkButton.Location = '192, 260'
$OkButton.Name = "OkButton"
$OkButton.Size = '75, 25'
$OkButton.TabIndex = 6
$OkButton.Text = "Ok"
$OkButton.FlatStyle = 'System'
$OkButton.ForeColor = 'Black'
$OkButton.add_Click($OkButton_OnClick)

$form.Controls.Add($OkButton)

$groupBox1.Location = '12, 165'
$groupBox1.Name = "groupBox1"
$groupBox1.Size = '576, 83'
$groupBox1.TabStop = $False
$groupBox1.Text = "Options:"

$form.Controls.Add($groupBox1)

$checkBox2.Location = '482, 15'
$checkBox2.Name = "checkBox2"
$checkBox2.Size = '88, 24'
$checkBox2.TabIndex = 5
$checkBox2.Text = "Verbose (/V)"
$checkBox2.UseVisualStyleBackColor = $True
$checkbox2.Checked = $true

$groupBox1.Controls.Add($checkBox2)

$checkBox1.Location = '97, 15'
$checkBox1.Name = "checkBox1"
$checkBox1.Size = '104, 24'
$checkBox1.TabIndex = 4
$checkBox1.Text = "Recursive (/S)"
$checkBox1.UseVisualStyleBackColor = $True
$checkbox1.Checked = $true

$groupBox1.Controls.Add($checkBox1)

$checkBox3.Location = '270, 15'
$checkBox3.Name = "checkBox3"
$checkBox3.Size = '130, 24'
$checkBox3.TabIndex = 6
$checkBox3.Text = "*.json Only (/Pattern)"
$checkBox3.UseVisualStyleBackColor = $True

$groupBox1.Controls.Add($checkBox3)

$label5.Location = '1, 46'
$label5.Name = "label5"
$label5.Size = '77, 23'
$label5.Text = "Log Location:"
$label5.TextAlign = 32

$groupBox1.Controls.Add($label5)

$button4.Location = '501, 47'
$button4.Name = "button4"
$button4.Size = '75, 23'
$button4.TabIndex = 4
$button4.Text = "Browse"
$button4.FlatStyle = 'System'
$button4.ForeColor = 'Black'
$button4.add_Click($button4_OnClick)

$groupBox1.Controls.Add($button4)

$textBox4.Location = '97, 50'
$textBox4.Name = "textBox4"
$textBox4.Size = '398, 20'
$textBox4.TabIndex = 3
$textBox4.Text = "ONLY CONSOLE LOGS FOR DOWNLOAD"

$groupBox1.Controls.Add($textBox4)

$label2.Location = '13, 126'
$label2.Name = "label2"
$label2.Size = '90, 23'
$label2.Text = "Destination:"
$label2.TextAlign = 16

$form.Controls.Add($label2)

$BrowseButton.Location = '513, 83'
$BrowseButton.Name = "BrowseButton"
$BrowseButton.Size = '75, 25'
$BrowseButton.TabIndex = 0
$BrowseButton.Text = "Browse"
$BrowseButton.FlatStyle = 'System'
$BrowseButton.ForeColor = 'Black'
$BrowseButton.add_Click($BrowseButton_OnClick)

$DestList=[collections.arraylist]@(
[pscustomobject]@{Name='PRODUCT FOLDER 1';Value="D:\TFS\Hyve Product Integration\Trunk\Products\PRODUCT FOLDER 1\"}
[pscustomobject]@{Name='PRODUCT FOLDER 2';Value="D:\TFS\Hyve Product Integration\Trunk\Products\PRODUCT FOLDER 2\"}
[pscustomobject]@{Name='PRODUCT FOLDER 3';Value="D:\TFS\Hyve Product Integration\Trunk\Products\PRODUCT FOLDER 3\"}
)
$DropDownBox = New-Object System.Windows.Forms.ComboBox 
$DropDownBox.Location = New-Object System.Drawing.Size(109,126) 
$DropDownBox.Size = New-Object System.Drawing.Size(479,20) 
$DropDownBox.DropDownHeight = 200 
$Form.Controls.Add($DropDownBox) 
$DropDownBox.DataSource=$DestList
$DropDownBox.DisplayMember='Name'

$SourceList1=[collections.arraylist]@(
[pscustomobject]@{Name='PRODUCT FOLDER 1';Value="/Root/Hyve Product Integration/Trunk/Products/PRODUCT FOLDER 1/Manifests"}	
[pscustomobject]@{Name='PRODUCT FOLDER 2';Value="/Root/Hyve Product Integration/Trunk/Products/PRODUCT FOLDER 2/Manifests"}
[pscustomobject]@{Name='PRODUCT FOLDER 3';Value="/Root/Hyve Product Integration/Trunk/Products/PRODUCT FOLDER 3/Manifests"}
)
$DropDownBox1 = New-Object System.Windows.Forms.ComboBox 
$DropDownBox1.Location = New-Object System.Drawing.Size(109,86) 
$DropDownBox1.Size = New-Object System.Drawing.Size(398,20) 
$DropDownBox1.DropDownHeight = 200 
$Form.Controls.Add($DropDownBox1) 
$DropDownBox1.DataSource=$SourceList1
$DropDownBox1.DisplayMember='Name'

$form.Controls.Add($BrowseButton)

$label1.Location = '12, 86'
$label1.Name = "label1"
$label1.Size = '78, 23'
$label1.Text = "Source:"
$label1.TextAlign = 16

$form.Controls.Add($label1)

$label4.Location = '176, 39'
$label4.Name = "label4"
$label4.Size = '225, 40'
$label4.TabIndex = 10
$label4.Text = "Andrews Super Manifest And Report Template Downloader"
$label4.Font = "Segoe UI, 10pt, style=Bold"
$label4.TextAlign = 'MiddleCenter'

$form.Controls.Add($label4)

$form.ResumeLayout()
$InitialFormWindowState = $form.WindowState
$form.add_Load($Form_StateCorrection_Load)
$form.add_FormClosed($Form_Cleanup_FormClosed)
$form.add_Shown({$form.Activate()})
return $form.ShowDialog()

} #End Function

Write-Host "Launching GUI..." -ForegroundColor Green		
Call-GUI | Out-Null

If ($EndScript -eq 1) {
	Write-Host "Script cancelled" -ForegroundColor Red
	Exit
	}
	
If (!$Target -or !$Source) {
	Write-Host "Please complete all fields and try again.." -ForegroundColor Red
	Sleep -Seconds 2
	Call-Gui
	}

$Options = [system.string]::Join(" ",$OptArray)


$Username = Read-Host "Enter Username"
$Domain = Read-Host "Domain"
$Password = Read-Host "Enter Password" -asSecureString

$TfsUrl = "https://TFSURL.COM/tfs/DefaultCollection"
$TeamProject = "GUID-57c3-4b8c-9a55-XXXXXGUID"
$ZipFileName = "TFS Download.zip"

[Reflection.Assembly]::LoadWithPartialName("System") | Out-Null
[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null

Write-Host "Checking whether download directory exists"
If(!(test-path $Target))
{
	Write-Host "Creating download directory"
    New-Item -ItemType Directory -Force -Path $Target
}

$Source = [System.Web.HttpUtility]::UrlEncode($Source)
$Url = $TfsUrl + "/" + $TeamProject + "/_api/_versionControl/itemContentZipped?repositoryId=&path=" + $Source


if([String]::IsNullOrEmpty($ZipFileName)){
     $index = $url.LastIndexOf("%2F") + 3
     $ZipFileName = $Url.Substring($index, ($Url.Length - $index))
     
}

Write-Host "Clearing download directory"
Get-ChildItem -Path $Target\\* -Recurse | Remove-Item -Force -Recurse
Get-ChildItem -Path "$Target" -Recurse | Remove-Item -Force -Recurse
$ZipFileName = Join-Path $Target $ZipFileName

$webClient = New-Object System.Net.WebClient
$webClient.Credentials = New-Object System.Net.NetworkCredential($Username, $Password,$Domain)

Write-Host "Downloading Zip"
$webClient.DownloadFile($Url, $ZipFileName)

$CopiedZip=Get-ChildItem $Target | select  -Last 1

Write-Host "Extracting Zip"
Get-ChildItem $Target -Filter *.zip | Expand-Archive -DestinationPath $Target -Force
$path = Join-Path $Target $CopiedZip

Write-Host "Removing Zip"
Remove-Item -Path $path