



<#
.Synopsis
   Gets a Lync/Skype For Business user's scheduled meeting URLs
.DESCRIPTION
   v1.0: 6 November 2015
   This is mostly useful after someone's SIP address has changed, but might be nice to run
   before, as well - there is no trace of the old meeting URLs after the SIP address change!

   Was originally written to complement Steve Parankewich's Exchange Web Service/PowerShell
   tool for finding and updating meeting invites after a SIP address change:

   http://powershellblogger.com/2015/10/automate-sip-address-and-upn-name-changes-in-lync-skype-for-business/

   You need to be running PowerShell as a user with access to the Front End server databases
   and connected to Lync/Skype for Business Management Shell.

.EXAMPLE
   Get-UpdatedConferenceURL -sipAddress 'sip:jane.smith@awesome.com'
.EXAMPLE
   Another example of how to use this cmdlet
   Get-CsUser -OU awesome.com/siteWeJustAcquired | Get-UpdatedConferenceURL
#>
function Get-UpdatedConferenceURL
{
    [CmdletBinding()]
    [OutputType([PSObject])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $sipAddress
    )

    Begin
    {
    }
    Process
    {
    # get user's current SIP address
        if ($sipAddress -notlike "sip:*") {
            $sipAddress = "sip:" + $sipAddress
        }
    # get CsUser object associated with SIP address
        $csUser = Get-CsUser $sipAddress
    # split prefix and domain
        $sipAddress = $sipAddress.Replace('sip:','')
        $sipPrefix = $sipAddress.split('@')[0]
        $sipDomain = $sipAddress.split('@')[1]
    # get simple URL for meetings
        $simpleUrl = ((Get-CsSimpleUrlConfiguration).SimpleUrl | where { $_.component -eq "Meet" -and $_.domain -eq $sipDomain }).ActiveUrl

    # find primary server that user is homed on
        $primaryUserServer = (Get-CsUserPoolInfo $csUser.identity).PrimaryPoolPrimaryRegistrar
    # get conferences from RTCLOCAL DB
        $conferences = Query-ConferenceDB -sipAddress $sipAddress -primaryUserServer $primaryUserServer
    # for each conference, output DB content, plus full meeting link
        foreach ($conference in $conferences) {
            $link = $simpleUrl + '/' + $sipPrefix + '/' + $conference.ExternalConfId
            $conference | Add-Member -MemberType NoteProperty -Name Link -Value $link
            $conference | Add-Member -MemberType NoteProperty -Name Prefix -Value $sipPrefix
            $conference
        }
 
    }
    End
    {
    }
}



# Helper function for querying the database

function Query-ConferenceDB {
Param($sipAddress
, $primaryUserServer)
    $database = 'RTC'
    $server = $primaryUserServer + '\RTCLOCAL'
    $sqlText = "SELECT [ConfId]
      ,[OrganizerId]
      ,[ServerMode]
      ,[Static]
      ,[ExternalConfId]
      ,[ExternalConfIdCksum]
      ,[ConferenceKey]
      ,[ConferenceKeyLax]
      ,[Title]
      ,[Description]
      ,[NotificationData]
	  ,CONVERT(varchar(4000), CAST([NotificationData] as varbinary(8000))) AS 'DecodedNotification'
      ,[OrganizerData]
      ,[ProvisionTime]
      ,[ExpiryTime]
      ,[LastUpdateTime]
      ,[LastActivationVersion]
      ,[LastActivateTime]
      ,[Version]
      ,[AdmissionType]
      ,[Locked]
      ,[Autopromote]
      ,[PstnLobbyBypass]
      ,[PstnAuthorityId]
      ,[PstnLocalId]
    FROM [rtc].[dbo].[Conference] c join [rtc].[dbo].[Resource] r on c.OrganizerId = r.ResourceId
    WHERE r.UserAtHost = '$sipAddress'"
    $connection = new-object System.Data.SqlClient.SQLConnection("Data Source=$server;Integrated Security=SSPI;Initial Catalog=$database");
    $cmd = new-object System.Data.SqlClient.SqlCommand($sqlText, $connection);

    $connection.Open();
    $reader = $cmd.ExecuteReader()

    $results = @()
    while ($reader.Read())
    {
        $row = @{}
        for ($i = 0; $i -lt $reader.FieldCount; $i++)
        {
            $row[$reader.GetName($i)] = $reader.GetValue($i)
        }
        $results += new-object psobject -property $row            
    }
    $connection.Close();

    $results
}
