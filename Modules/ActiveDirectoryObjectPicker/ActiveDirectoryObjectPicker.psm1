if ( $PSVersionTable.PSVersion -lt '2.0' ) {

    throw 'Unsupported PowerShell Version'

}

if ( $PSVersionTable.PSEdition -eq 'Core' ) {

    $DotNetCoreVersion = [System.Environment]::Version

    if ( $DotNetCoreVersion -gt '5.0' ) {

        $AssemblyFolder = 'net5.0-windows7.0'

    } elseif ( $DotNetCoreVersion -gt '3.1' ) {

        $AssemblyFolder = 'netcoreapp3.1'

    } elseif ( $DotNetCoreVersion -gt '3.0' ) {

        $AssemblyFolder = 'netcoreapp3.0'

    } else {

        throw 'Requires at least .NET Core v3.0'

    }

} elseif ( $PSVersionTable.CLRVersion -ge '4.0' ) {

    $DotNet45Release = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release |
        Select-Object -First 1 -ExpandProperty Release

    $DotNet4 = Get-ItemProperty -Path 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Client'
    
    if ( $DotNet45Release -ge 528040 ) {

        $AssemblyFolder = 'net48'

    } elseif ( $DotNet45Release -ge 461808 ) {

        $AssemblyFolder = 'net472'

    } elseif ( $DotNet45Release -ge 394802 ) {

        $AssemblyFolder = 'net462'

    } elseif ( $DotNet45Release -ge 379893 ) {

        $AssemblyFolder = 'net452'

    } elseif ( $DotNet45Release -ge 378389 ) {

        $AssemblyFolder = 'net45'

    } else {

        $AssemblyFolder = 'net40'

    }

} elseif ( $PSVersionTable.CLRVersion -ge '2.0' ) {

    [bool]$DotNet35Installed = Get-ItemProperty -Path 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.5' -Name Install |
        Select-Object -ExpandProperty Install

    if ( $DotNet35Installed ) {

        $AssemblyFolder = 'net35'

    } else {

        $AssemblyFolder = 'net20'

    }

}

$Assembly = Get-Item -Path "$PSScriptRoot\Tulpep.ActiveDirectoryObjectPicker*\lib\$AssemblyFolder\ActiveDirectoryObjectPicker.dll" |
    Select-Object -Last 1

Add-Type -Path $Assembly.FullName

<#
.SYNOPSIS
 Create an Active Directory object picker dialog object

.DESCRIPTION
 Create an Active Directory object picker dialog object using the Tulpep.ActiveDirectoryObjectPicker assembly

.LINK
 https://github.com/Tulpep/Active-Directory-Object-Picker

.PARAMETER AllowedLocations
 The scopes the dialog is allowed to search.

.PARAMETER AllowedObjectTypes
 A list of LDAP attribute names that will be retrieved for picked objects.

.PARAMETER AttributesToFetch
 Which attributes should be returned along with the object

.PARAMETER DefaultLocations
 The initially selected scope in the dialog.

.PARAMETER DefaultObjectTypes
 The initially selected types of objects in the dialog.

.PARAMETER MultiSelect
 Whether the user can select multiple objects.

.PARAMETER Providers
 The providers affecting the ADPath returned in objects.

.PARAMETER ShowAdvancedView
 Gets or sets whether objects flagged as show in advanced view only are
 displayed (up-level).

.PARAMETER SkipDomainControllerCheck
 Gets or sets the whether to check whether the target is a Domain Controller
 and hide the "Local Computer" scope.

 If this flag is NOT set, then the DSOP_SCOPE_TYPE_TARGET_COMPUTER flag will be
 ignored if the target computer is a DC. This flag has no effect unless
 DSOP_SCOPE_TYPE_TARGET_COMPUTER is specified.

.PARAMETER TargetComputer
 Gets or sets the name of the target computer.

.PARAMETER Credential
 Use this method to override the user credentials, passing new credentials for the account profile to be used.
#>
function New-ActiveDirectoryObjectPicker {

    param(

        [Tulpep.ActiveDirectoryObjectPicker.Locations[]]
        $AllowedLocations = 'All',

        [Tulpep.ActiveDirectoryObjectPicker.ObjectTypes[]]
        $AllowedObjectTypes = 'All',

        [string[]]
        $AttributesToFetch = @(),

        [Tulpep.ActiveDirectoryObjectPicker.Locations[]]
        $DefaultLocations = 'JoinedDomain',

        [Tulpep.ActiveDirectoryObjectPicker.ObjectTypes[]]
        $DefaultObjectTypes = 'BuiltInGroups,Groups,Users',

        [switch]
        $MultiSelect,

        [Tulpep.ActiveDirectoryObjectPicker.ADsPathsProviders[]]
        $Providers = 'Default',

        [switch]
        $ShowAdvancedView,

        [switch]
        $SkipDomainControllerCheck,

        [string]
        $TargetComputer,

        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty

    )

    $Picker = New-Object -TypeName Tulpep.ActiveDirectoryObjectPicker.DirectoryObjectPickerDialog
    
    $Picker.AllowedLocations          = $AllowedLocations
    $Picker.AllowedObjectTypes        = $AllowedObjectTypes
    $Picker.AttributesToFetch         = $AttributesToFetch
    $Picker.DefaultLocations          = $DefaultLocations
    $Picker.DefaultObjectTypes        = $DefaultObjectTypes
    $Picker.MultiSelect               = $MultiSelect.IsPresent
    $Picker.Providers                 = $Providers
    $Picker.ShowAdvancedView          = $ShowAdvancedView.IsPresent
    $Picker.SkipDomainControllerCheck = $SkipDomainControllerCheck.IsPresent
    $Picker.TargetComputer            = $TargetComputer
    $Picker.Tag                       = $Tag
    $Picker.Site                      = $Site

    if ( $Credential -ne [System.Management.Automation.PSCredential]::Empty ) {

        $Picker.SetCredentials( $Credential.UserName, $Credential.GetNetworkCredential().Password )

    }

    return $Picker

}

<#
.SYNOPSIS
 Show an Active Directory object picker dialog and return the results

.DESCRIPTION
 Show an Active Directory object picker dialog and return the results

.LINK
 https://github.com/Tulpep/Active-Directory-Object-Picker

.PARAMETER AllowedLocations
 The scopes the dialog is allowed to search.

.PARAMETER AllowedObjectTypes
 A list of LDAP attribute names that will be retrieved for picked objects.

.PARAMETER AttributesToFetch
 Which attributes should be returned along with the object

.PARAMETER DefaultLocations
 The initially selected scope in the dialog.

.PARAMETER DefaultObjectTypes
 The initially selected types of objects in the dialog.

.PARAMETER MultiSelect
 Whether the user can select multiple objects.

.PARAMETER Providers
 The providers affecting the ADPath returned in objects.

.PARAMETER ShowAdvancedView
 Gets or sets whether objects flagged as show in advanced view only are
 displayed (up-level).

.PARAMETER SkipDomainControllerCheck
 Gets or sets the whether to check whether the target is a Domain Controller
 and hide the "Local Computer" scope.

 If this flag is NOT set, then the DSOP_SCOPE_TYPE_TARGET_COMPUTER flag will be
 ignored if the target computer is a DC. This flag has no effect unless
 DSOP_SCOPE_TYPE_TARGET_COMPUTER is specified.

.PARAMETER TargetComputer
 Gets or sets the name of the target computer.

.PARAMETER Credential
 Use this method to override the user credentials, passing new credentials for the account profile to be used.
#>
function Show-ActiveDirectoryObjectPicker {

    param(

        [Tulpep.ActiveDirectoryObjectPicker.Locations[]]
        $AllowedLocations,

        [Tulpep.ActiveDirectoryObjectPicker.ObjectTypes[]]
        $AllowedObjectTypes,

        [string[]]
        $AttributesToFetch,

        [Tulpep.ActiveDirectoryObjectPicker.Locations[]]
        $DefaultLocations,

        [Tulpep.ActiveDirectoryObjectPicker.ObjectTypes[]]
        $DefaultObjectTypes,

        [switch]
        $MultiSelect,

        [Tulpep.ActiveDirectoryObjectPicker.ADsPathsProviders[]]
        $Providers,

        [switch]
        $ShowAdvancedView,

        [switch]
        $SkipDomainControllerCheck,

        [string]
        $TargetComputer,

        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty

    )

    $Picker = New-ActiveDirectoryObjectPicker @PSBoundParameters
    $Button = $Picker.ShowDialog()
    $Picker.Dispose()

    if ( $Button -eq 'Cancel' ) { return }

    return $Picker.SelectedObjects

}
