function Get-O365LastPasswordChangeTimeStamp{
[cmdletbinding()]
    param (      
        [parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Please enter a username or list of usernames.")]
                   [ValidateNotNullOrEmpty()]
                   [string[]]$username,
        
        [parameter(Mandatory=$true,
                   HelpMessage="Please provide a crednetial obejct")]
                   [ValidateNotNullOrEmpty()]
                   [System.Management.Automation.CredentialAttribute()]$credential
        ) 
    
  
    <#
    .SYNOPSIS 
    Retrieves UserPricipalName, DisplayName, and LastPasswordChangeTimeStamp attributes from Office 365 user

    .DESCRIPTION
    This function will query against office 365 to get the last time a users password has been changed.
    Takes input as a list or single user.  Additionally, your office 365 crednetials must be passed for this function to work.

    .PARAMETER username
    Specify a single or a list of usernames.  These can either be username or username@contoso.com

    .PARAMETER Credential
    Specifices a set of credentials used to query data from Office 365

    .INPUTS
    You can pipe a txt file of usernames to this function.
   
    .EXAMPLE
    C:\PS> Get-O365LastPasswordChangeTimeStamp -username username@contoso.com -credential $cred

    .EXAMPLE
    C:\PS> Get-content C:\users\username\Desktop\usernamelist.txt | Get-O365LastPasswordChangeTimeStamp -credential $cred

    #>
    
BEGIN{
    $accountdata =@()
    write-host "Checking to see if the Microsoft Online PowerShell module is installed"
    if ((get-module -name MSOnline -ErrorAction SilentlyContinue | foreach { $_.Name }) -ne "MSOnline"){
        write-host "Microsoft Online Management PowerShell is not added to this session, adding it now..."
        try{
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $credential -Authentication Basic -AllowRedirection
        Import-PSSession $Session
        Import-Module MSOnline
        Connect-MsolService -Credential $credential
        }
        catch{
            write-host "Please make sure you have the 'Microsoft Online Services Sign-In Assistant for IT professionals RTW' installed: http://www.microsoft.com/en-us/download/details.aspx?id=41950 `n
                        After installing the Sing-In Assistant, please install the 'Windows Azure Active Directory Module for Windows PowerShell (64-bit version)': http://go.microsoft.com/fwlink/p/?linkid=236297"
            break
        }
    }
    else{
        write-host Microsoft Online PowerShell module is good to go. -backgroundcolor black -foregroundcolor green
        start-sleep -s 1
        Connect-MsolService -Credential $credential
    }

}

PROCESS{
    foreach ($name in $username){
        $accountdata += Get-MsolUser -UserPrincipalName $name | select UserPrincipalName,DisplayName,LastPasswordChangeTimeStamp
    }

}
END{
    return $accountdata
}
}
