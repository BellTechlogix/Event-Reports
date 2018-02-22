<#
	ServerReboots-Reports.ps1
	Created By Kristopher Roy
	Created On 21Feb18
#>

#Functions Section
#This Function creates a dialogue to return a Folder Path
	function Get-Folder {
		param([string]$Description="Select Folder to place results in",[string]$RootFolder="Desktop")

	 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
		 Out-Null     

	   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
			$objForm.Rootfolder = $RootFolder
			$objForm.Description = $Description
			$Show = $objForm.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
			If ($Show -eq "OK")
			{
				Return $objForm.SelectedPath
			}
			Else
			{
				Write-Error "Operation cancelled by user."
			}
	}

	#File Select Function - Lets you select your input file
	function Get-FileName
	{
	  param(
		  [Parameter(Mandatory=$false)]
		  [string] $Filter,
		  [Parameter(Mandatory=$false)]
		  [switch]$Obj,
		  [Parameter(Mandatory=$False)]
		  [string]$Title = "Select A File",
		  [Parameter(Mandatory=$False)]
		  [string]$InitialDirectory
		)
		if(!($Title)) { $Title="Select Input File"}
		if(!($InitialDirectory)) { $InitialDirectory="c:\"}
		[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
		$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
		$OpenFileDialog.initialDirectory = $initialDirectory
		$OpenFileDialog.FileName = $Title
		#can be set to filter file types
		IF($Filter -ne $null){
			$FilterString = '{0} (*.{1})|*.{1}' -f $Filter.ToUpper(), $Filter
			$OpenFileDialog.filter = $FilterString}
		if(!($Filter)) { $Filter = "All Files (*.*)| *.*"
			$OpenFileDialog.filter = $Filter}
		$OpenFileDialog.ShowDialog() | Out-Null
		IF($OBJ){
			$fileobject = GI -Path $OpenFileDialog.FileName.tostring()
			Return $fileObject}
		else{Return $OpenFileDialog.FileName}
	}


	#This function lets you build an array of specific list items you wish
	#This function lets you build an array of specific list items you wish
	Function MultipleSelectionBox ($inputarray,$prompt,$listboxtype,$label) {
 
	# Taken from Technet - http://technet.microsoft.com/en-us/library/ff730950.aspx
	# This version has been updated to work with Powershell v3.0.
	# Had to replace $x with $Script:x throughout the function to make it work. 
	# This specifies the scope of the X variable.  Not sure why this is needed for v3.
	# http://social.technet.microsoft.com/Forums/en-SG/winserverpowershell/thread/bc95fb6c-c583-47c3-94c1-f0d3abe1fafc
	#
	# Function has 3 inputs:
	#     $inputarray = Array of values to be shown in the list box.
	#     $prompt = The title of the list box
	#     $listboxtype = system.windows.forms.selectionmode (None, One, MultiSimple, or MultiExtended)
 
	$Script:x = @()
 
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
 
	$objForm = New-Object System.Windows.Forms.Form 
	$objForm.Text = $prompt
	$objForm.Size = New-Object System.Drawing.Size(300,600) 
	$objForm.StartPosition = "CenterScreen"
 
	$objForm.KeyPreview = $True
 
	$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
		{
			foreach ($objItem in $objListbox.SelectedItems)
				{$Script:x += $objItem}
			$objForm.Close()
		}
		})
 
	$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
		{$objForm.Close()}})
 
	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Size(75,520)
	$OKButton.Size = New-Object System.Drawing.Size(75,23)
	$OKButton.Text = "OK"
 
	$OKButton.Add_Click(
	   {
			foreach ($objItem in $objListbox.SelectedItems)
				{$Script:x += $objItem}
			$objForm.Close()
	   })
 
	$objForm.Controls.Add($OKButton)
 
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(150,520)
	$CancelButton.Size = New-Object System.Drawing.Size(75,23)
	$CancelButton.Text = "Cancel"
	$CancelButton.Add_Click({$objForm.Close()})
	$objForm.Controls.Add($CancelButton)
 
	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(10,20) 
	$objLabel.Size = New-Object System.Drawing.Size(280,20) 
	$objLabel.Text = $label
	if($objLabel.Text -eq $null -or $objLabel.Text -eq ""){$objLabel.Text = "Please make a selection from the list below:"}
	$objForm.Controls.Add($objLabel) 
 
	$objListbox = New-Object System.Windows.Forms.Listbox 
	$objListbox.Location = New-Object System.Drawing.Size(10,40) 
	$objListbox.Size = New-Object System.Drawing.Size(260,20) 
 
	$objListbox.SelectionMode = $listboxtype
 
	$inputarray | ForEach-Object {[void] $objListbox.Items.Add($_)}
 
	$objListbox.Height = 470
	$objForm.Controls.Add($objListbox) 
	$objForm.Topmost = $True
 
	$objForm.Add_Shown({$objForm.Activate()})
	[void] $objForm.ShowDialog()
 
	Return $Script:x
	}

#set initial script functionality options
$outputoptions = "Screen","File"
$OutPutSelection = MultipleSelectionBox -inputarray $outputoptions -listboxtype MultiSimple -label "Where Do You Want Your Report?:" -prompt "OutPut Selection"

$getoptions = "AutoList","ManualInput"
$GetSelection = MultipleSelectionBox -inputarray $getoptions -listboxtype One -label "Choose to Return an Auto List or Manual Input:" -prompt "Server Input Mode"

#OutPut Selection Functionality
IF($OutPutSelection -eq "File")
{
    #Hide PowerShell Window
    $sig='[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
    Add-Type -MemberDefinition $sig -name NativeMethods -namespace Win32
    $PSId= @(Get-Process|where{$_.ProcessName -like "powershell*"} -ErrorAction SilentlyContinue)[0].MainWindowHandle
    If ($PSId -ne $NULL) { [Win32.NativeMethods]::ShowWindowAsync($PSId,2)}

    #Get Output Folder Selection
    write-host "Select the folder to place the results in, may be hidden behind PowerShell window"
    $folder = get-folder
    
    #Restore PowerShell Window
    $sig='[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
    Add-Type -MemberDefinition $sig -name NativeMethods -namespace Win32
    $PSId= @(Get-Process|where{$_.ProcessName -like "powershell*"} -ErrorAction SilentlyContinue)[0].MainWindowHandle
    If ($PSId -ne $NULL) { [Win32.NativeMethods]::ShowWindowAsync($PSId,1)}

    write-host "folder selected $folder"
}

#Get a list of Servers
##All Servers
IF($GetSelection -eq "AutoList")
{
	$serverlist = get-adcomputer -Filter{OperatingSystem -like "*Server*"}
	#Build your serverlist prompt box
	$autolist = $serverlist.name|Sort-Object
	$ServerSelections = MultipleSelectionBox -listboxtype multisimple -inputarray $autolist
}
##Manual Selection of Servers
IF($GetSelection -eq "ManualInput")
{
	$manualoptions = "Typed","CSV/List"
	$manualselection = MultipleSelectionBox -inputarray $manualoptions -listboxtype One -label "How Do You Wish to Input Servers:" -prompt "Manual Input Mode"
	IF($manualselection -eq "Typed")
	{
		[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
		$title = 'Servers'
		$msg   = 'Enter your Server List Comma Seperated:'
		$ServerList = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
        $Serverselections = $ServerList.Split(",")
	}
	IF($manualselection -eq "CSV/List")
	{$ServerSelections = (Import-Csv (Get-FileName) -header "Servers").servers}
}

write-host $ServerSelections

$reportarray=@()

Foreach($server in $serverSelections)
{
    $details = gwmi win32_ntlogevent -filter "LogFile='System' and EventCode='1074' and Message like '%restart%'" -ComputerName $server | select User,@{n="Time";e={$_.ConvertToDateTime($_.TimeGenerated)}},Message
    FOREACH($item in $details|where{$_ -ne $null})
    {
        $rebootobject = new-object PSObject
        $item
        $rebootobject | Add-Member NoteProperty -Name "Server" -Value $Server
        $rebootobject | Add-Member NoteProperty -Name "User" -Value $item.User
        $rebootobject | Add-Member NoteProperty -Name "Time" -Value $item.Time
        $rebootobject | Add-Member NoteProperty -Name "Message" -Value $item.Message
        $reportarray += $rebootobject
        $rebootobject = $null
        $item = $null
    }
}

$date = Get-Date -Format ddMMMyy-HHmm
If($folder -ne $null)
{
    $reportarray|export-csv $folder"\serverreboots-$date.csv" -NoTypeInformation
    $folder = $null
}

If($OutPutSelection -eq "Screen"){$reportarray|ft}