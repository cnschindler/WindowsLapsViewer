Active Directory Object Picker for PowerShell
========================
### The standard Active Directory object picker dialog for .NET, now as a PowerShell module!

This project is based on the [Active Directory Object Picker](https://github.com/Tulpep/Active-Directory-Object-Picker) maintained by [Tulpep](https://github.com/Tulpep), which is in turn based on [Active Directory Common Dialogs .NET (ADUI)](https://adui.codeplex.com/) created in 2004 by Armand du Plessis.

It has been wrapped in a few function definitions and packaged as a PowerShell module for everyday use.

### How to use it
You can install the latest version from PSGallery
```powershell
Install-Module ActiveDirectoryObjectPicker
```

And use it this way:
```powershell
PS> Show-ActiveDirectoryObjectPicker

FetchedAttributes : {}
Name              : Ford Prefect
Path              : LDAP://domain.local/CN=Ford Prefect,CN=Users,DC=domain,DC=local
SchemaClassName   : user
Upn               : ford.prefect@domain.local

```
