function Search-O365MessageFromIP{
[cmdletbinding()]
    param (      
        [parameter(Mandatory=$true,
                   ParameterSetName='ip',
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Please enter a FromIP you are searching for.  FromIP is the 'Sent From' IP address of a message.")]
                   [AllowEmptyString()]
                   [string[]]$FromIP,

        [parameter(ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Please enter a StartDate for your message search.  If not provided, search will be 30 days.")]
                   [AllowEmptyString()]
                   [switch[]]$StartDate,

        [parameter(ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Please enter a StartDate for your message search.  If not provided, search will be today's date.")]
                   [AllowEmptyString()]
                   [switch[]]$EndDate,

        [parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Please provide a crednetial obejct")]
                   [ValidateNotNullOrEmpty()]
                   [System.Management.Automation.CredentialAttribute()]$credential
        ) 

    <#
    .SYNOPSIS 
    Get message details from a specific O365 user

    .DESCRIPTION
    This function will receive message details for a specific user for the last 48 hours, unless ranage (up to 30 days) is specified with 
    additional parameters (StartDate & EndDate).  This function will return the "From IP" of each message listed within Exchange online.
    Takes input as a list or single user.  Additionally, your office 365 crednetials must be passed for this function to work.

    .PARAMETER username
    Specify a single or a list of usernames.  These can either be username or username@contoso.com

    .PARAMETER FromIP
    Specify the From IP you are wanting to search messages for. This will collect all messages set within a specific time frame sent from a specific IP.

    .PARAMETER Credential
    Specifices a set of credentials used to query data from Office 365

    .INPUTS
    You can pipe a txt file of usernames to this function.
   
    .EXAMPLE
    C:\PS> Search-O365MessageFromIP -FromIP "41.8.8.4" -credential $cred

    .EXAMPLE
    C:\PS> Search-O365MessageFromIP -FromIP "41.8.8.4" -StartDate $(Get-Date -format d) -EndDate $(Get-Date).AddDays(-29).ToString("d") -credential $cred

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

    if ($StartDate){
        foreach ($ip in $FromIP){
            $fromIpData += Get-MessageTrace -FromIP $FromIP -StartDate $StartDate -EndDate $EndDate
        }
    }
    else{
        foreach ($ip in $FromIP){
            $fromIpData += Get-MessageTrace -FromIP $FromIP
        }
    }
}
END{
    return $fromIpData
}
}