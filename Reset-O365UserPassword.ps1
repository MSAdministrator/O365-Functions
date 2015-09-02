function Reset-O365UserPassword{
[cmdletbinding()]
    param (      
        [parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Please enter a username or list of usernames.")]
                   [ValidateNotNullOrEmpty()]
                   [string[]]$username,
        [parameter(ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Please select if you want the user to change their password the next time they log into their account.")]
                   [AllowEmptyString()]
                   [switch[]]$changePasswordAtNextLogin,

        [parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Please provide a crednetial obejct")]
                   [ValidateNotNullOrEmpty()]
                   [System.Management.Automation.CredentialAttribute()]$credential
        ) 
    
  
    <#
    .SYNOPSIS 
    Reset's the password for the specified Office 365 user account

    .DESCRIPTION
    This function will reset an office 365 users password.
    Takes input as a list or single user.  Additionally, your office 365 crednetials must be passed for this function to work.

    .PARAMETER username
    Specify a single or a list of usernames.  These can either be username or username@contoso.com

    .PARAMETER changePasswordAtNextLogin
    Specify whether or not the user should change their password (after the reset) at next login.

    .PARAMETER Credential
    Specifices a set of credentials used to query data from Office 365

    .INPUTS
    You can pipe a txt file of usernames to this function.
   
    .EXAMPLE
    C:\PS> Get-O365LastPasswordChangeTimeStamp -username username@contoso.com -changePasswordAtNextLogin $false -credential $cred

    .EXAMPLE
    C:\PS> Get-content C:\users\username\Desktop\usernamelist.txt | Get-O365LastPasswordChangeTimeStamp -changePasswordAtNextLogin $true -credential $cred

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

    if ($changePasswordAtNextLogin){
        foreach ($name in $username){
            $accountdata += Set-MsolUserPassword -UserPrincipalName $name -ForceChangePassword $true
        }
    }
    else{
        foreach ($name in $username){
            $accountdata += Set-MsolUserPassword -UserPrincipalName $name -ForceChangePassword $false
        }
    }
}
END{
    return $accountdata
}
}