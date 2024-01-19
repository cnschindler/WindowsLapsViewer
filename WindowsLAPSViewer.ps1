#
# WindowsLapsViewer by Christian Schindler, NTx BOCG, christian.schinder@ntx.at
#
# Latest Version at "https://github.com/cnschindler/WindowsLapsViewer"
#
# Provided as is. No liabilty.
#

[cmdletbinding(SupportsShouldProcess = $true)]
Param()


$ScriptParentPath = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
$ScriptModulePath = Join-Path -Path $ScriptParentPath -ChildPath "Modules"

#Region Logging
[string]$Script:LogFileNamePrefix = "NTxWindowsLapsViewer"
[string]$Script:LogfileName = ($LogFileNamePrefix + "_{0:yyyyMMdd-HHmmss}.log" -f [DateTime]::Now)
[string]$Script:LogPath = $ScriptParentPath
[string]$script:LogFilePath = Join-Path -Path $Script:LogPath -ChildPath $Script:LogfileName    
[string]$Script:LogFileStart = "Logging started"
[string]$Script:LogFileStop = "Logging stopped"
$Script:NoLogging = $true
#EndRegion Logging

function Write-LogFile
{
    # Logging function, used for progress and error logging...
    # Uses the globally (script scoped) configured variables 'LogFilePath' to identify the logfile and 'NoLogging' to disable it.
    #
    [CmdLetBinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$LogPrefix,
        [System.Management.Automation.ErrorRecord]$ErrorInfo = $null
    )

    # Prefix the string to write with the current Date and Time, add error message if present...
    if ($ErrorInfo)
    {
        $logLine = "{0:d.M.y H:mm:ss} : {1}: {2} Error: {3}" -f [DateTime]::Now, $LogPrefix, $Message, $ErrorInfo.Exception.Message
    }

    Else
    {
        $logLine = "{0:d.M.y H:mm:ss} : {1}: {2}" -f [DateTime]::Now, $LogPrefix, $Message
    }

    if (-not $NoLogging)
    {
        # Create the Script:Logfile and folder structure if it doesn't exist
        if (-not (Test-Path $Script:LogFilePath -PathType Leaf))
        {
            New-Item -ItemType File -Path $Script:LogFilePath -Force -Confirm:$false -WhatIf:$false | Out-Null
            #Add-Content -Value "Logging started." -Path $Script:LogFilePath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        }

        # Write to the Script:Logfile
        Add-Content -Value $logLine -Path $Script:LogFilePath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        Write-Verbose $logLine
    }
    
    Else
    {
        Write-Host $logLine
    }
}

Function ManageModuleAndSnapInLoading
{
    # Function to check for and import modules

    Param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [switch]$Snapin,
        [switch]$IsFile
    )

    if ($Snapin)
    {
        if ((Get-PSSnapin $Name -Registered -ErrorAction SilentlyContinue))
        {
            try
            {
                Add-PSSnapin -Name $Name -ErrorAction Stop
                Write-LogFile -Message "$($Name)$($LoadSnapinSuccess)"
            }
            
            catch
            {
                Write-LogFile -Message "$($LoadSnapinError)$($Name)." -ErrorInfo $_
            }
        }

        Else
        {
            Write-LogFile -Message "$($EX_RecipientMGMTSnapin_NotInstalled)"
        }
    }

    Else
    {
        $IsModuleInstalled = (Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1)
    
        if ($IsFile)
        {
            [System.IO.FileInfo]$Filename = $Name
            $MatchedName = $Filename.Name.Replace(".psm1", "")
        }

        Else
        {
            $MatchedName = $Name
        }

        if ($IsModuleInstalled.Name -match "$($MatchedName)")
        {   
            try
            {
                Import-Module -Name $Name -ErrorAction Stop -WarningAction SilentlyContinue -DisableNameChecking
                Write-LogFile -Message "$($Name)$($LoadModuleSuccess)"
            }
            
            catch
            {
                Write-LogFile -Message "$($LoadModuleError)$($Name)." -ErrorInfo $_
            }
        }
    
        Else
        {
            Write-LogFile = "$($Name)$($ModuleNotInstalled)."
            Break
        }    
    }
}
function LoadFileBasedModules
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]
        $Path
    )

    $ModuleFiles = Get-ChildItem -Path $Path -Recurse -Filter "*.psm1"

    foreach ($File in $ModuleFiles)
    {
        ManageModuleAndSnapInLoading -Name $file.Fullname -IsFile
    }
}
Function Manage-Modules
{
    # Function to check for and import modules

    Param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    $IsModuleInstalled = (Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1)
    
    if ($IsModuleInstalled.Name -eq "$($ModuleName)")
    {   
        try
        {
            Import-Module -Name $ModuleName -ErrorAction Stop -WarningAction SilentlyContinue -DisableNameChecking
            $Textbox_Messages.Text = "LAPS Module successfully loaded!"
        }
        
        catch
        {
            $Textbox_Messages.Text = "LAPS Module could not be loaded. Error: $($Error.Exception.InnerException)"
            $Button_RetrievePassword.Visibility = 1
        }
    }

    else
    {
        $Textbox_Messages.Text = "LAPS Module not installed. Please install first!"
        $Button_RetrievePassword.Visibility = 1
    }
}
function Get-LAPSClearTextPassword
{
    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        $Computer
    )

    Try
    {
        $ComputerClearTextPasswords = Get-LapsADPassword -Identity $Computer -AsPlainText -IncludeHistory -ErrorAction Stop | Select-Object Password, PasswordUpdateTime
        
        If (-not $ComputerClearTextPasswords)
        {
            $Textbox_Messages.Text = "No Passwords could be retrieved. Check your permissions!"
        }

        else
        {
            $Textbox_Messages.text = "Successfully retrieved passwords"
            Return $ComputerClearTextPasswords
        }
    }

    Catch
    {
        $Textbox_Messages.Text = $_
    }
}

function Get-ObjectPickerSelection
{
    
        $Attributes = "SamAccountName"
        $AllowedObjectTypes = "Computers"        

    $ReturnObject = Show-ActiveDirectoryObjectPicker -AttributesToFetch $Attributes -AllowedObjectTypes $AllowedObjectTypes
    if ($ReturnObject.Name -match ".")
    {
        Return $ReturnObject            
    }

    else
    {
        Show-MessageBox -BoxTitle $MSGBoxTitleInvalidSelection -BoxMessageText  $MessageNoObjectSelected -BoxIcon $MSGBoxWarn -Buttons $MSGBoxOK    
    }
 
}

#Region XAML Form
Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms, WindowsFormsIntegration

[xml]$XAMLForm = @'

<Window x:Name="Windows_Laps_Viewer" x:Class="WindowsLapsViewer.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WindowsLapsViewer"
        mc:Ignorable="d"
        Title="NTx Windows Active Directory LAPS Viewer" Width="530" MinWidth="530" MinHeight="380" ResizeMode="NoResize" ScrollViewer.VerticalScrollBarVisibility="Disabled" SizeToContent="Height">
    <Window.Resources>
        <Style x:Key="Datagrid" TargetType="{x:Type DataGrid}">
            <Setter Property="Background" Value="White"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="SelectionMode" Value="Single"/>
            <Setter Property="SelectionUnit" Value="Cell"/>
            <Setter Property="CanUserResizeColumns" Value="False"/>
            <Setter Property="HorizontalScrollBarVisibility" Value="Hidden"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
        </Style>
    </Window.Resources>
    <Grid VerticalAlignment="Top" HorizontalAlignment="Left" Width="530" Height="380">
        <Label x:Name="Label_Computername" Content="Computername" HorizontalAlignment="Left" Margin="20,10,0,0" VerticalAlignment="Top" Width="99" FontWeight="Bold"/>
        <TextBox x:Name="Textbox_Computername" HorizontalAlignment="Left" Height="20" Margin="20,35,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="255" TabIndex="1" FontSize="12"/>
        <Button x:Name="Button_RetrievePassword" HorizontalAlignment="Right" Height="40" Margin="0,25,60,0" VerticalAlignment="Top" Width="80">
            <AccessText Text="Retrieve Passwords" TextWrapping="Wrap" TextAlignment="Center"/>
        </Button>
        <Button x:Name="Button_SelectComputer" HorizontalAlignment="Left" Height="40" Margin="290,25,0,0" VerticalAlignment="Top" Width="80">
            <AccessText Text="Select Computer..." TextWrapping="Wrap" TextAlignment="Center"/>
        </Button>
        <Label x:Name="Label_LapsPassword" Content="Current LAPS Password" HorizontalAlignment="Left" Height="27" Margin="20,60,0,0" VerticalAlignment="Top" Width="150" FontWeight="Bold"/>
        <TextBox x:Name="Textbox_Lapspassword" HorizontalAlignment="Left" Height="20" Margin="20,85,0,0" VerticalAlignment="Top" Width="470" FontSize="12"/>
        <Label x:Name="Label_LapsHistory" Content="LAPS History entries" HorizontalAlignment="Left" Height="27" Margin="20,110,0,0" VerticalAlignment="Top" Width="137" FontWeight="Bold"/>
        <DataGrid x:Name="Datagrid_LapsHistory" Margin="20,135,0,0" AutoGenerateColumns="False" HorizontalAlignment="Left" VerticalAlignment="Top" Height="110" Width="470">
            <DataGrid.ContextMenu>
                <ContextMenu>
                    <MenuItem Name="Datagrid_CopyContextMenu" Header="Copy"/>
                </ContextMenu>
            </DataGrid.ContextMenu>
            <DataGrid.Resources>
                <Style TargetType="DataGridCell">
                    <Setter Property="Margin" Value="0"/>
                    <Setter Property="Padding" Value="0"/>
                    <Setter Property="BorderThickness" Value="0"/>
                </Style>
            </DataGrid.Resources>
            <DataGrid.Columns>
                <DataGridTextColumn Header="Password" Binding="{Binding Password}" Width="405" />
                <DataGridTextColumn Header="Date Set" Binding="{Binding PasswordUpdateTime}" Width="80" />
            </DataGrid.Columns>
        </DataGrid>
        <Label x:Name="Label_Messages" Content="Messages" HorizontalAlignment="Left" Height="27" Margin="20,260,0,0" VerticalAlignment="Top" Width="137" FontWeight="Bold"/>
        <TextBox x:Name="Textbox_Messages" HorizontalAlignment="Left" Height="20" Margin="20,285,0,0" VerticalAlignment="Top" Width="470" FontSize="12"/>
        <TextBlock x:Name="Textblock_Info" HorizontalAlignment="Left" Margin="20,340,0,0" TextWrapping="Wrap" Text="Windows LAPS Viewer by Christian Schindler, NTx BOCG, christian.schindler@ntx.at" VerticalAlignment="Top" Width="482"/>
    </Grid>
</Window>

'@ -replace 'mc:Ignorable="d"', '' -replace "x:Name", 'Name' -replace '^<Win.*', '<Window' -replace 'x:Class="\S+"', ''

# Read XAMLForm
$form = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XAMLForm))
$XAMLForm.SelectNodes("//*[@Name]") | Where-Object { Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name) }

# Icon definition
#
# Icon in Base64 format
[string]$IconB64 = @"
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
"@

# Convert icon to bitmap
$bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
$bitmap.BeginInit()
$bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($IconB64)
$bitmap.EndInit()
$bitmap.Freeze()

# Add icon to form
$Form.Icon = $bitmap
#endregion

# Configure LAPS History Datagrid
$Datagrid_LapsHistory.IsReadOnly = $true
$Datagrid_LapsHistory.SelectionMode = "Single"
$Datagrid_LapsHistory.SelectionUnit = "Cell"
$Datagrid_LapsHistory.CanUserResizeColumns = $False
$Datagrid_LapsHistory.HorizontalScrollBarVisibility = "Hidden"
$Datagrid_LapsHistory.HeadersVisibility = "Column"
$Datagrid_LapsHistory.GridLinesVisibility = "None"
$Datagrid_LapsHistory.Background = "White"

# Disable Retrieve Passwort Button by default
$Button_RetrievePassword.IsEnabled = $false

# Set Focus on the Computername Textbox
$Textbox_Computername.Focus() | Out-Null

# Handler for changed text in the Textbox
$Textbox_Computername.Add_TextChanged(
    {
        $Textbox_Messages.Clear()
        $Textbox_Lapspassword.Clear()
        $Button_RetrievePassword.IsEnabled = $true
    }
)

# Handler for Retrieve Password Button click
$Button_RetrievePassword.Add_Click(
    {
        $Textbox_Messages.Clear()
        $Datagrid_LapsHistory.items.Clear()
        $Computername = $Textbox_Computername.Text

        if (-not $Computername)
        {
            $Textbox_Messages.Text = "No Computername was specified!"
        }

        else
        {
            $RetrievedPasswords = Get-LAPSClearTextPassword -Computer $Computername
            $Textbox_Lapspassword.Text = $RetrievedPasswords[0].Password

            if ($RetrievedPasswords.Count -gt 1)
            {
                Foreach ($item in $RetrievedPasswords)
                {
                    if ($RetrievedPasswords.IndexOf($item) -eq 0)
                    {
                        Continue
                    }

                    else
                    {
                        $UpdateTime = $item.PasswordUpdateTime.ToShortDateString()
                        $Datagrid_LapsHistory.AddChild([pscustomobject]@{Password = $item.Password; PasswordUpdateTime = $UpdateTime })
                    }
                }
            }
        }
    }    
)

# Handler for Select Computer Button click
$Button_SelectComputer.Add_Click(
    {
        $Computer = Get-ObjectPickerSelection
        if ($Computer.Name -match ".")
        {
            $Textbox_Computername.Text = $Computer.FetchedAttributes[0]
        }
    }
)

# Handler for datagrid contextmenu
$Datagrid_CopyContextMenu.Add_Click({
        $Item = $Datagrid_LapsHistory.CurrentItem
        try
        {
            $item.Password | Set-Clipboard -ErrorAction Stop
            $Textbox_Messages.Text = "Password copied to clipboard"
        }

        Catch
        {
            $Textbox_Messages.Text = "Error copying password to clipboard. Error: $($Error.Exception.InnerException)"
        }
        
    }
)

# Import Windows LAPS Module
ManageModuleAndSnapInLoading -Name ActiveDirectory | Out-Null
#Manage-Modules -ModuleName LAPS | Out-Null
LoadFileBasedModules -Path $ScriptModulePath

# Load Form
$Form.ShowDialog() | Out-Null
