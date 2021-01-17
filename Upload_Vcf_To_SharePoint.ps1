#region help

<#  
.SYNOPSIS
    Upload vcf to Sharepoint Online Business Card Document Libary.

.DESCRIPTON
    1) Connect to SharePoint Online
    2) Upload each .vcf file in C:\<your filepath here> to Business Card Document Libary.

.NOTES
    Created by: Josh Crabtree 18 Jan 2021
#>

#endregion help

#region variables

# Cred Setup
$Username = "<YOUR USERNAME HERE>"
$AesKey = Get-Content "C:\<PATH TO AES KEY >"
$Password = Get-Content "C:\<PATH TO PASSWORD>" | ConvertTo-SecureString -Key $AesKey
$Creds = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $Password

#dates & log file
$Today = (Get-Date).ToString('MM-dd-yyyy')
$logfile = "C:\<PATH TO LOG FILE HERE>-$($Today).log"

#SharePoint Site & Library
$SPOSite = "https://<PATH TO YOUR SHAREPOINT LIBRARY SITE HERE>"
$SPOLibrary = "Business Cards"

#endregion variables

#Import SharePointPnpModule
Import-Module SharePointPnPPowerShellOnline

#region functions

#region Log-It Function
#'Log-It' Function used for color-coded screen output and output to the log file.
function Log-It {
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $True,
            Position = 0,
            ValueFromPipeline=$True
        )]
        [String]$Message,
        [ValidateSet(
            "General","Process","Success","Failure","Warning","Notification","LogOnly","ScreenOnly"
        )]
        [String]$Status = "General"
    )
    Switch($Status){
        "General"{
            $Color="Cyan"
            $Type="[INFORMA] "
        }
        "Process"{
            $Color="White"
            $Type="[PROCESS] "
        }
        "Failure"{
            $Color="Red"
            $Type="[FAILURE] "
        }
        "Success"{
            $Color="Green"
            $Type="[SUCCESS] "
        }
        "Warning"{
            $Color="Yellow"
            $Type="[WARNING] "
        }
        "Notification"{
            $Color="Gray"
            $Type="[NOTICES] "
        }
        "ScreenOnly"{
            $Color="Magenta"
            $Type="[INFORMA] "
        }
        "LogOnly"{
            $Color=$Null
            $Type="[INFORMA] "
        }

    }
    if($Color -ne $Null){Write-Host -ForegroundColor $Color $Type$Message}
    if($Color -ne "Magenta"){"$($(Get-Date).ToString('yyyy-MM-dd::hh:mm:ss')) | $Type$Message" | Out-File $logfile -Append}
}
#endregion Log-It Function

#region Connect-SPPnPOnline Function
#connect to SharePoint Online using the PnP Module
function Connect-SPPnPOnline{
    $FN = "Connect SharePoint PnP Online"
    
    try{
        Connect-PnPOnline -Url $sposite -Credential $creds
        "$FN | Connected to SharePoint PnP Online" | Log-it -status Success
    }
    
    catch{
        "$FN | Failed to connect to SharePoint PnP Online | $($error[0].Exception.Message)" | Log-it -status Failure
    }
}
#endregion Connect-SPPnPOnline Function

#region Upload-VcfToSharePointLibrary Function
#Connect to the SharePoint library defined in the variable above and upload the .vcf file to the library
function Upload-VcfToSharePointLibrary{
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true,
            position = 0,
            ValueFromPipeline = $true
        )]
        $file
    )
    begin{
         $FN = "Upload-PhotoToSharePointLibrary"
        "$FN | BEGIN: Upload-PhotoToSharePointLibrary Functon." | Log-It -Status Notification
    }
    process{
        try{
            Add-PnPFile -Folder $SpoLibrary -Path $File.FullName
            "$FN | Added file $($file.fullname) to $SpoLibrary" | Log-it -Status Success
        }
        catch{
            "$FN | Failed to add $($file.fullname) to $SpoLibrary |$($error[0].Exception.Message)" | Log-it -Status Failure
        }
    }
    end{
        "$FN | END: Upload-PhotoToSharePointLibrary Functon." | Log-It -Status Notification
    }
}
#endregion Upload-VcfToSharePointLibrary Function

#endregion functions

#region processing
Connect-SPPnPOnline

#get all the files in the directory ending in .vcf
$VcfDir = Get-ChildItem -Path C:\<YOUR FILE PATH HERE>\* -Include *.vcf

#foreach file in the directory, upload the .vcf files to the SharePoint library defined in your variables
foreach($vcf in $VcfDir){
    $vcf | Upload-VcfToSharePointLibrary
}
#endregion processing