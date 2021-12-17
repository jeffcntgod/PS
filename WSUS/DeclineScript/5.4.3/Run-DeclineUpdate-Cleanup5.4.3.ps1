
<#PSScriptInfo

.VERSION 5.4.2

.GUID f201f241-04ae-409a-9038-e44b51cd5769

.AUTHOR Jeff Carreon (twiiter: @jeffctangsoo10) or (email: jeffcntgod@gmail.com)

.COMPANYNAME Keep Calm Carre-on!

.COPYRIGHT 2021 Keep Calm Carre-on!. All rights reserved.

.TAGS 

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES

	$DeclineLastLevelOnly & $trial CANNOT be used at the same time.
	Script will create custom application "CMSDKPosh" Eventlog for storing events below
		EventID 21020 = Successful Run
		EventID 21021 = Error running the main function
		EventID 21031 = Error running either the Decline superseded or itanium function


	Version:	5.4.3
	Author:		Jeff Carreon

    Updates: ver. 5.4.3  (12/17/2021)
        - Added Metadatasize comparison report (Requires Importing SQLPS or SQLServer Module)

    Updates: ver. 5.4.2  (12/15/2021)
        - Added a function to decline Windows 10 Feature Updates for Enablement package
        - Added -forcesync.  For forcing SUP synchronization from Top down on CM Hierarchies
        - Fixed the CleanUpdateList function
        
    Updates: ver. 5.4.1  (3/9/2021)
        - Updated the report to not show categories with 0 results.  Though it will list the ones with 0 below the table.

	Updates: ver. 5.4  (3/2/2021)
		- Added an OneOff Manual Decline function.  For declining single patches or multiple depending on the -kb input (below)
        - example usage:  .\Run-DeclineUpdate-Cleanup.ps1 -trialrun -OneOffCleanup -kb "*KB2768005*"
        NOTEs: 
            - The -OneOffCleanup depends on -kb being populated
            - The -kb uses the "like" operator. 
            - I strongly recommened using the -Trialrun first, then validate the list of patches that are documented in the logs and html it creates before declining.

	Updates: ver. 5.3  (12/9/2020)
		- Added a function for Windows 10 versions 1507/1511/1607/1703/1803/1903/2004  
        - Added a function for declining legacy Office and M365

    Updates: ver. 5.2  (10/14/2020)
        - Added the following, but only -CleanupObsoleteComputers is being used.
        	Invoke-WsusServerCleanup -CompressUpdates
	        Invoke-WsusServerCleanup -CleanupObsoleteComputers
	        Invoke-WsusServerCleanup -CleanupUnneededContentFiles
	
	Updates: ver. 5.1 (8/7/2018)
		- Fixed the missing comma in one of the paramaters (Thanks Johan for pointing that out!)
        - Improved/Updates OS filtering, to only allow decline of targetted OS/Updates.
        - Added Decline updates for Windows 7, Windows 8, Windows 8.1, Windows Server 2003, Windows Server 08, Windows Server 08 R2, Windows Server 12, and Windows Server 12 R2.  
		NOTE: All are SKIPPED by default. Modify the corresponding skip statement in the parameter secion below accordingly and remove the $true, if you would like to decline any of these.

	Updates 5/10/2018: 
		- Added Decline updates for ARM64-Based, and IE 10
        - Added Clean Update List maintnance function (optional), deletes files/folders that are # of days old.
		- Fixed error handling on querying for updates.
		- Perfomance improvement

	Updates 4/25/2018: 
		- Added Decline updates for Win10 Next, and Server Next
        - Added email reporting and logging
		- Perfomance improvement on querying updates

#>

<# 
.SYNOPSIS
	Script is for declining superseeded, Itanium, Preview, Beta, ARM64, IE7, IE8, IE9, IE10, Win10 Next, Server Next, Embedded, Legacy Win10, Unwanted M365 updates, and various Legacy Windows OS (See below) Updates in WSUS/SUP environment.
	    Windows 7, Windows 8, Windows 8.1, Windows Server 2003, Windows Server 08, Windows Server 08 R2, Windows Server 12, and Windows Server 12 R2.

    It can also now target and decline individual patches...
	
    Recommend running this monthly...  
	Run the scripts targetting the bottom or downstreams servers (bottom SUPs), then run it against the upstream server (Top SUP)...
	BE AWARE:  
		The OS filtering function may grab Other articles. I HIGHLY recommend doing a -TrialRun first, and examine the results in the 'UpdatesList' folder before executing this in prod environment.  
			To do this, either run the script with -TrialRun or set this switch to $true (which is on by default), see param section below.  
			AND remove the = $true off the OS/products' switches below if you'd like to see list of the updates that you'd like to decline first, before setting the -TrialRun to false.

.DESCRIPTION 
 Script is designed to decline all of the updates that have been superseded for over 90 days (by default), and MORE!!!. 


	# $Servers				= Specify the target servers as default target(s) for automation.  Or, can be specified manually at run time.
	# $UseSSL               = Specify whether WSUS Server is configured to use SSL
	# $Port                 = Specify WSUS Server Port (Hard coded in param section, though this can be specified otherwise)
	# $TrialRun		        = Specify this to do a test run and get a summary of how many superseded updates, Itanium Updates, XP, Preview, Beta, Win10 Next, Server Next, ARM64, IE7, IE8, IE9, IE10, and Embedded Updates there are that can be declined.  It records to .csv and htm file
	# $DeclineLastLevelOnly = Specify whether to decline all superseded updates or only last level superseded updates
									**For example, Update1 supersedes Update2. Update2 supersedes Update3. In this scenario, the Last Level in the supersedence chain is Update3. 
									**To decline only the last level updates in the supersedence chain, specify the DeclineLastLevelOnly switch
	# $ExclusionPeriod      = Specify the number of days between today and the release date for which the superseded updates must not be declined. Eg, if you want to keep superseded updates published within the last 2 months, specify a value of 60 (days)
	# $SkipItanium			= Specify this or set to true, to skip declining Itanium updates.
	# $SkipXP				= Specify this or set to true, to skip declining Windows XP updates.  
	# $SkipPrev				= Specify this or set to true, to skip declining Windows Preview updates.  
	# $SkipBeta				= Specify this or set to true, to skip declining Windows Beta updates.
	# $SkipWin10Next		= Specify this or set to true, to skip declining Windows 10 Next updates.
	# $SkipWin7				= Specify this or set to true, to skip declining Windows 7 updates. Default is $true
	# $SkipWin8				= Specify this or set to true, to skip declining Windows 8 updates. Default is $true
	# $SkipWin81			= Specify this or set to true, to skip declining Windows 8.1 updates. Default is $true
	# $SkipWin2k3			= Specify this or set to true, to skip declining Windows Server 2003 updates. Default is $true
	# $SkipWin2k8			= Specify this or set to true, to skip declining Windows Server 2008 updates. Default is $true
	# $SkipWin2k8R2			= Specify this or set to true, to skip declining Windows Server 2008 R2 updates. Default is $true
	# $SkipWin12			= Specify this or set to true, to skip declining Windows Server 12 updates. Default is $true
	# $SkipWin12R2			= Specify this or set to true, to skip declining Windows Server 12 R2 updates. Default is $true
	# $SkipServerNext		= Specify this or set to true, to skip declining Windows Server Next updates.
	# $SkipArm64			= Specify this or set to true, to skip declining ARM64-based updates.
	# $SkipIE7				= Specify this or set to true, to skip declining IE7 updates.
	# $SkipIE8				= Specify this or set to true, to skip declining IE8 updates. Default is $true
	# $SkipIE9				= Specify this or set to true, to skip declining IE9 updates. Default is $true
	# $SkipIE10				= Specify this or set to true, to skip declining IE10 updates. Default is $true
	# $SkipEmbedded				= Specify this or set to true, to skip declining Windows Embedded updates.

	# $SkipLegacyWin10			= Specify this or set to true, to skip declining Legacy Win10 Upgrades.

						- Current Criteria
							($_.Title -match "Windows 10 Version 1507" -and $_.Title -notmatch "Server") -or
							($_.Title -match "Windows 10 Version 1511" -and $_.Title -notmatch "Server") -or 
							($_.Title -match "Windows 10 Version 1607" -and $_.Title -notmatch "Server") -or 
							($_.Title -match "Windows 10 Version 1703" -and $_.Title -notmatch "Server") -or 
							($_.Title -match "Windows 10 Version 1803" -and $_.Title -notmatch "Server") -or 
							($_.Title -match "Windows 10 Version 1903" -and $_.Title -notmatch "Server") -or 
							($_.Title -match "Windows 10 Version 2004" -and $_.Title -notmatch "Server")
                                     
	 # $SkipLegacyOff365			= Specify this or set to true, to skip declining unwanted Office and M365 updates.

						- Current Criteria
							($_.Title -match "Office 365" -and $_.Title -match "x86 based Edition") -or 
							($_.Title -match "Microsoft 365 Apps Update" -and $_.Title -match "x86 based Edition")
    
	# $SkipWin10FeatureUpdates  = Specify this or set to true, to skip declining unwanted Windows 10 Feature Updates / Enablement Packages
    
						-Current Criteria
							($_.UpdateClassificationTitle -eq "Upgrades") -and !(($_.Title -match "x64-based") -and ($_.Title -match "Enablement Package"))


	# $CompressUpdates			= If true, this runs Invoke-WsusServerCleanup -CompressUpdates.  Default is false

	# $CleanupObsoleteComputers	= If true, this runs Invoke-WsusServerCleanup -CleanupObsoleteComputers.  Default is true

	# $CleanupUnneededContentFiles = If true, this runs Invoke-WsusServerCleanup -CleanupUnneededContentFiles.  Default is false
    
	# $OneOffCleanup & $kb		= Specify this or set to true, to skip declining unwanted Office and M365 updates.c


	# $CleanUpdatelist			= Specify whether to clean the UpdateList folders/files to prevent build up. Default is $true
	# $CleanULNumber			= Specify the number of days old folders/files to keep in UpdateList folder

	# $forcesync			= Specify whether SUP sync should be run, after all servers have done declines.  Default is $false

    # $MetaDataReportingOn	= Specify whether Metadata size comparison report should be on or not. Default is $false (NOTE:  This requires SQLServer or SQLPS Module to be imported for Invoke-SQLCMD to work)

#> 

[CmdletBinding()]
Param(
	
	# Define lower tier SUP servers first, then TOP WSUS/SUP server last in this array
	$Servers = @("server1.domain.com","server2.domain.com","server3.domain.com","TopSup.domain.com"),	

	[bool]$UseSSL = $false,
	
	[int]$PortNumber = 8530,
	
	[switch] $TrialRun,

	[string] $CMprovider = "<CMServerProvider>",

	[string] $SiteCode = "<SITECODE>",

    [string] $SUSDBServer = "SUSDB SERVER (Top SUP)",
	
	[switch] $DeclineLastLevelOnly,
	
	[Parameter(Mandatory=$False)]
	[int] $ExclusionPeriod = 90,

	[switch] $SkipItanium,
	
	[switch] $SkipXP,
	
	[switch] $SkipPrev = $true,
	
	[switch] $SkipBeta = $true,

	[switch] $SkipWin10Next,

	[switch] $SkipWin7,

	[switch] $SkipWin8,

	[switch] $SkipWin81,

	[switch] $SkipWin2k3,

	[switch] $SkipWin2k8,
	
	[switch] $SkipWin2k8R2,

	[switch] $SkipWin12,
	
	[switch] $SkipWin12R2 = $true,

	[switch] $SkipServerNext,
	
	[switch] $SkipIE7,
	
	[switch] $SkipIE8,
	
	[switch] $SkipIE9,
	
	[switch] $SkipIE10,
	
	[switch] $SkipEmbedded,

	[switch] $SkipArm64,

	[switch] $SkipLegacyWin10,
    
	[switch] $SkipLegacyOff365,

	[switch] $SkipWin10FeatureUpdates,

	[switch] $CompressUpdates = $false,

	[switch] $CleanupObsoleteComputers = $true,

	[switch] $CleanupUnneededContentFiles = $false,
	
	[switch] $OneOffCleanup,

	[string] $kb,

	[switch] $forcesync = $true,

	[switch] $MetaDataReportingOn = $true,

	[bool]$EmailReport = $true,
	
	[string]$SMTPServer = "relayserver.domain.com",
	
	[string]$From = "yourteam@mail.com",
	
	[string[]]$To = "dude1@mail.com,dude2@mail.com,dude3@mail.com",

	[string]$Subject = "WSUS/SUP Decline Updates Report",

	[string]$ReportTitle = "WSUS/SUP Decline Updates Maintenance Task",
	
	# UpdateList Folder maintenance
	[switch] $CleanUpdatelist = $true,
	
	# Define # of days old before it Cleans Update list files and folders
	[int]$CleanULNumber = 90

)

$error.Clear()
$SinglePatch = $kb
$script:zerovals = New-Object System.Collections.ArrayList
$ScriptVersion = "5.4.3"
$EventSource = "WSUS Decline Maintenance"
$Eventlog = "CMSDKPosh"
$td = (get-date -uformat %m-%d-%y)
$path = Get-Location
$scriptName = $MyInvocation.MyCommand.Name
$ul = "UpdatesList"
$ulpath = "$path\" + "$ul\" + "$td"
$delpath = "$path\" + "$ul"
[int]$xcounter = 0
[int]$scount= $Servers.count
$Overallhtmfile = "$ulpath\" + "_OverallCountsSummary-$td.htm"

$CStyle = "<Style>BODY{font-size:12px;font-family:verdana,sans-serif;color:navy;font-weight:normal;}" + `
			"TABLE{border-width:1px;cellpadding=10;border-style:solid;border-color:navy;border-collapse:collapse;}" + `
			"TH{font-size:12px;border-width:1px;padding:10px;border-style:solid;border-color:navy;}" + `
			"TD{font-size:10px;border-width:1px;padding:10px;border-style:solid;border-color:navy;}</Style>"



[String]$LogFile = "$path\" + $($((Split-Path $MyInvocation.MyCommand.Definition -leaf)).replace("ps1","log")) #Name and Location of LogFile


if ([System.Diagnostics.EventLog]::SourceExists('WSUS Decline Maintenance') -ne "True")
{
    New-EventLog -LogName $Eventlog -Source $EventSource
}

 Write-EventLog -LogName $Eventlog -EventID 21020 -Message "Run-DeclineCleanup Script has started." -Source $EventSource -EntryType Information

If($TrialRun){$Subject += " Trial Run"}
Function SendEmailStatus($From, $To, $Subject, $SMTPServer, $BodyAsHtml, $Body)
{	
    $SMTPMessage = New-Object System.Net.Mail.MailMessage $From, $To, $Subject, $Body
	$SMTPMessage.IsBodyHTML = $BodyAsHtml
	$SMTPClient = New-Object Net.Mail.SMTPClient($SMTPServer)
    $SMTPClient.Send($SMTPMessage)
	If($? -eq $False){Write-Warning "$($Error[0].Exception.Message) | $($Error[0].Exception.GetBaseException().Message)"}
	$SMTPMessage.Dispose()
	rv SMTPClient
	rv SMTPMessage
}

Function Write-toFile{
    <#
    .SYNOPSIS
        Writing information to file
    .DESCRIPTION
        Function to write information to file
    #>    
    Param ([string]$WriteLine)
    Out-File $XMLFile -encoding utf8 -input $WriteLine -append   
    Write-Host $WriteLine
}

Function Write-ToLog([string]$message, [string]$file) {
    <#
    .SYNOPSIS
        Writing log to the logfile
    .DESCRIPTION
        Function to write logging to a logfile. This should be done in the End phase of the script.
    #>
    If(-not($file)){$file=$LogFile}        
    $Date = $(get-date -uformat %Y-%m-%d-%H.%M.%S)
    $message = "$Date | `t$message"
    Write-Verbose $message
    Write-Host $message
    #Write Log to log file Without ASCII not able to read with tracer.
    Out-File $file -encoding ASCII -input $message -append
}

Function GetSuperSededList {

	$Script:countAllUpdates = 0
	$Script:countSupersededAll = 0
	$Script:countSupersededLastLevel = 0
	$Script:countSupersededExclusionPeriod = 0
	$Script:countSupersededLastLevelExclusionPeriod = 0
	$Script:countDeclined = 0
	
    $Prop = [ordered]@{}
    $ErrorActionPreference = "Stop"
    foreach($update in $allUpdates)
    {
    
        $Script:countAllUpdates++
    
        if ($update.IsDeclined) {
            $Script:countDeclined++
        }
    
        if (!$update.IsDeclined -and $update.IsSuperseded) {
            $Script:countSupersededAll++
        
            if (!$update.HasSupersededUpdates) {
                $Script:countSupersededLastLevel++
            }
			###################
            if ($update.CreationDate -lt (get-date).AddDays(-$ExclusionPeriod))  {
				#"$($update.Id.UpdateId.Guid), $($update.Id.RevisionNumber), $($update.Title), $($update.KnowledgeBaseArticles), $($update.SecurityBulletins), $($update.HasSupersededUpdates)" | Out-File $outSupersededExList -Append 
                "$($update.Title), $($update.KnowledgeBaseArticles), $($update.ArrivalDate), $($update.SecurityBulletins), $($update.UpdateClassificationTitle), $($update.ProductTitles), $($update.HasSupersededUpdates)" | Out-File $outSupersededExList -Append
		        $Script:countSupersededExclusionPeriod++
			    if (!$update.HasSupersededUpdates) {
				    $Script:countSupersededLastLevelExclusionPeriod++
			    }
            }		
        
            "$($update.Id.UpdateId.Guid), $($update.Id.RevisionNumber), $($update.Title), $($update.KnowledgeBaseArticles), $($update.ArrivalDate), $($update.SecurityBulletins), $($update.HasSupersededUpdates)" | Out-File $outSupersededList -Append       
        
            $Prop.Title = [string]$update.Title
            $Prop."KB Article" = [string]$update.KnowledgeBaseArticles
            $Prop."Arrival Date" = [string]$update.ArrivalDate
            $Prop.Classification = [string]$update.UpdateClassificationTitle
            $Prop."Product Title" = [string]$update.ProductTitles
            $Prop."Has Superseded Updates" = [string]$update.HasSupersededUpdates
               
            New-Object PSObject -property $Prop
        }
     }
     
}

Function Decline-Superseded{
	#Write-ToLog ""
    
    Write-ToLog ""
    Write-ToLog "$script:WsusServer is starting Decline-SupersededUpdates function..."

	"UpdateID, RevisionNumber, Title, KBArticle, ArrivalDate, SecurityBulletin, LastLevel" | Out-File $outSupersededList
    "Title, KBArticle, ArrivalDate, SecurityBulletin, UpdateClassificationTitle, ProductTitles, HasSupersededUpdates" | Out-File $outSupersededExList


##########################

    GetSuperSededList | ConvertTo-HTML -head $CStyle | Out-File $outSupersededHTM
    

	Write-ToLog "Done."
	if ($csv){Write-ToLog "List of superseded updates: $outSupersededList"}
	Write-ToLog "List of Superseded Updates: $outSupersededHTM"

	Write-ToLog ""
	Write-ToLog "Superseded Summary:"
	Write-ToLog "========"
	Write-ToLog "All Updates = $countAllUpdates"
	Write-ToLog "Any except Declined = ($countAllUpdates - $countDeclined)"
	Write-ToLog "All Superseded Updates = $countSupersededAll"
	Write-ToLog "    Superseded Updates (Intermediate) = ($countSupersededAll - $countSupersededLastLevel)"
	Write-ToLog "    Superseded Updates (Last Level) = $countSupersededLastLevel"
	Write-ToLog "    Superseded Updates (Older than $ExclusionPeriod days) = $countSupersededExclusionPeriod"
	Write-ToLog "    Superseded Updates (Last Level Older than $ExclusionPeriod days) = $countSupersededLastLevelExclusionPeriod"
	Write-ToLog ""

	$i = 0
	if (!$TrialRun) {
	    
	    Write-ToLog "TrialRun flag is set to $TrialRun. Continuing with declining updates"
	    $updatesDeclined = 0
	    
	    if ($DeclineLastLevelOnly) {
	        Write-ToLog "  DeclineLastLevel is set to True. Only declining last level superseded updates." 
	        
	        foreach ($update in $allUpdates) {
	            
	            if (!$update.IsDeclined -and $update.IsSuperseded -and !$update.HasSupersededUpdates) {
	              if ($update.CreationDate -lt (get-date).AddDays(-$ExclusionPeriod))  {
				    $i++
					$percentComplete = "{0:N2}" -f (($updatesDeclined/$countSupersededLastLevelExclusionPeriod) * 100)
					Write-Progress -Activity "Declining Updates" -Status "Declining update #$i/$countSupersededLastLevelExclusionPeriod - $($update.Id.UpdateId.Guid)" -PercentComplete $percentComplete -CurrentOperation "$($percentComplete)% complete"
					
	                try 
	                {
	                    $update.Decline()                    
	                    $updatesDeclined++
	                }
	                catch [System.Exception]
	                {
	                    Write-ToLog "$script:WsusServer failed to decline update $($update.Id.UpdateId.Guid). Error:" $_.Exception.Message
	                    Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer failed to decline update $($update.Id.UpdateId.Guid). Error: $error[0]" -Source $EventSource -EntryType Error
	                } 
	              }             
	            }
	        }        
	    }
	    else {
	        Write-ToLog "  DeclineLastLevel is set to False. Declining all superseded updates."
	        
	        foreach ($update in $allUpdates) {
	            
	            if (!$update.IsDeclined -and $update.IsSuperseded) {
	              if ($update.CreationDate -lt (get-date).AddDays(-$ExclusionPeriod))  {   
				  	
					$i++
					$percentComplete = "{0:N2}" -f (($updatesDeclined/$countSupersededAll) * 100)
					Write-Progress -Activity "Declining Updates" -Status "Declining update #$i/$countSupersededAll - $($update.Id.UpdateId.Guid)" -PercentComplete $percentComplete -CurrentOperation "$($percentComplete)% complete"
	                try 
	                {
	                    $update.Decline()
	                    $updatesDeclined++
	                }
	                catch [System.Exception]
	                {
	                    Write-Host "$script:WsusServer failed to decline update $($update.Id.UpdateId.Guid). Error:" $_.Exception.Message
	                    Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer failed to decline update $($update.Id.UpdateId.Guid). Error: $error[0]" -Source $EventSource -EntryType Error
	                }
	              }              
	            }
	        }   
	        
	    }
	    
	    Write-ToLog "  Declined $updatesDeclined updates."
	    #if ($updatesDeclined -ne 0) {
	    #    Copy-Item -Path $outSupersededList -Destination $outSupersededListBackup -Force
		#	Write-ToLog "  Backed up list of superseded updates to $outSupersededListBackup"
	    #}
	    
	}
	else {
	    Write-ToLog "TrialRun flag is set to $TrialRun. Skipped declining updates"
	
    if(Test-Path $outSupersededExList){ Import-csv $outSupersededExList | ConvertTo-HTML -head $table | Out-File $outSupersededExHTM}

    }
    
}



Function GetMetaDataSizeBefore{

	$MetaDataSizeBefore = Invoke-SQLcmd -serverInstance $SUSDBServer -Database "SUSDB" -Query "

    ;with cte as
    (
        SELECT dbo.tbXml.RevisionID, ISNULL(datalength(dbo.tbXml.RootElementXml), 0) as Uncompressed, ISNULL(datalength(dbo.tbXml.RootElementXmlCompressed), 0) as Compressed FROM dbo.tbXml
        INNER  JOIN dbo.tbProperty ON dbo.tbXml.RevisionID = dbo.tbProperty.RevisionID
    )

    Select
      Count(Distinct(u.LocalUpdateId)) As [Update Count],
      SUM(cte.Uncompressed) /1048576 as [Catalog Size (MB)],
      SUM(cte.Compressed) /1048576 as [Compressed Catalog Size (MB)]
    From tbUpdate u
      inner join tbRevision r on u.LocalUpdateID = r.LocalUpdateID
      inner join tbProperty pr on pr.RevisionID = r.RevisionID
      inner join cte on cte.revisionid = r.revisionid
    where r.RevisionID in
    (
        Select  t1.RevisionID
        From tbBundleAll t1
        Inner Join tbBundleAtLeastOne t2
        On t1.BundledID=t2.BundledID
        Where ishidden=0 and pr.ExplicitlyDeployable=1
    )

    " -ErrorVariable err1 -ErrorAction SilentlyContinue 

	
	
	[int]$script:SbeforeUpdateCount = $MetaDataSizeBefore.'Update Count'
	[int]$Script:SbeforeCatSize = $MetaDataSizeBefore.'Catalog Size (MB)'
	[int]$Script:SbeforeCompCatSize = $MetaDataSizeBefore.'Compressed Catalog Size (MB)'

}

Function GetMetaDataSizeAfter{

	$MetaDataSizeAfter = Invoke-SQLcmd -serverInstance $SUSDBServer -Database "SUSDB" -Query "

    ;with cte as
    (
        SELECT dbo.tbXml.RevisionID, ISNULL(datalength(dbo.tbXml.RootElementXml), 0) as Uncompressed, ISNULL(datalength(dbo.tbXml.RootElementXmlCompressed), 0) as Compressed FROM dbo.tbXml
        INNER  JOIN dbo.tbProperty ON dbo.tbXml.RevisionID = dbo.tbProperty.RevisionID
    )

    Select
      Count(Distinct(u.LocalUpdateId)) As [Update Count],
      SUM(cte.Uncompressed) /1048576 as [Catalog Size (MB)],
      SUM(cte.Compressed) /1048576 as [Compressed Catalog Size (MB)]
    From tbUpdate u
      inner join tbRevision r on u.LocalUpdateID = r.LocalUpdateID
      inner join tbProperty pr on pr.RevisionID = r.RevisionID
      inner join cte on cte.revisionid = r.revisionid
    where r.RevisionID in
    (
        Select  t1.RevisionID
        From tbBundleAll t1
        Inner Join tbBundleAtLeastOne t2
        On t1.BundledID=t2.BundledID
        Where ishidden=0 and pr.ExplicitlyDeployable=1
    )

    " -ErrorVariable err1 -ErrorAction SilentlyContinue 

	
	
	[int]$script:SAfterUpdateCount = $MetaDataSizeAfter.'Update Count'
	[int]$script:SAfterCatSize = $MetaDataSizeAfter.'Catalog Size (MB)'
	[int]$script:SAfterCompCatSize = $MetaDataSizeAfter.'Compressed Catalog Size (MB)'
}




Function Decline-Itanium{


	Write-ToLog "$script:WsusServer is starting Decline-WsusItaniumUpdates function..... Please wait...."
    #Write-ToLog "Connecting to $script:WsusServer..."

    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Itanium updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}


    Write-ToLog "Searching for Itanium updates..."
	$ItaniumUpdates = $GrabUpdates | where-object {$_.Title -match "ia64|itanium"}
	$script:Itancount = $ItaniumUpdates.count
	
	If($ItaniumUpdates)
	{
		Write-ToLog "Found $script:Itancount Itanium Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Itanium Updates...";$ItaniumUpdates | %{$_.Decline()}}Else{Write-ToLog "Recording Itanium Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $ItaniumUpdates | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $IThtmfile
            If(!$TrialRun){Write-ToLog "List of Itanium updates declined: $IThtmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Itanium updates that could be declined: $IThtmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Itanium Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Itanium Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Itanium Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-XPUpdates{



	Write-ToLog "$script:WsusServer is starting Decline-XPUpdates function..... Please wait...."
    #Write-ToLog "Connecting to $script:WsusServer..."



    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Windows XP updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Windows XP updates..."
	$XPUpdates = $GrabUpdates | where-object {$_.Title -match "Windows XP" -and $_.Title -notmatch "Server"}
	$Script:XPcount = $XPUpdates.count	
	
	If($XPUpdates)
	{
		Write-ToLog "Found $Script:XPcount Windows XP Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows XP Updates...";$XPUpdates | %{$_.Decline()}}Else{Write-ToLog "Recording Windows XP Updates..."}
            $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $XPUpdates | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $XPhtmfile
            If(!$TrialRun){Write-ToLog "List of Windows XP updates declined: $XPhtmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Windows XP updates that could be declined: $XPhtmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Windows XP Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Windows XP Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows XP Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Preview{

	Write-ToLog "$script:WsusServer is starting Decline-Preview function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Preview updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Preview updates..."
	$Prev = $GrabUpdates | where-object {$_.Title -match "Preview"}
	$Script:Prevcount = $Prev.count	
	
	If($Prev)
	{
		Write-ToLog "Found $Script:Prevcount Preview Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Preview Updates...";$Prev | %{$_.Decline()}}Else{Write-ToLog "Recording Preview Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $Prev | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Prevhtmfile
            If(!$TrialRun){Write-ToLog "List of Preview updates declined: $Prevhtmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Preview updates that could be declined: $Prevhtmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Preview Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Preview Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Preview Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Beta{
	

	Write-ToLog "$script:WsusServer is starting Decline-Beta function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Beta updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Beta updates..."
	$Beta = $GrabUpdates | where-object {$_.Title -match "Beta"}
	$Script:Betacount = $Beta.count	
	
	If($Beta)
	{
		Write-ToLog "Found $Script:Betacount Beta Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Beta Updates...";$Beta | %{$_.Decline()}}Else{Write-ToLog "Recording Beta Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $Beta | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Betahtmfile
            If(!$TrialRun){Write-ToLog "List of Beta updates declined: $Betahtmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Beta updates that could be declined: $Betahtmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Beta Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Beta Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Beta Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Win10Next{
	
	Write-ToLog "$script:WsusServer is starting Decline-Win10Next function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Windows 10 Next updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Windows 10 Next updates..."
	$Win10Next = $GrabUpdates | where-object {$_.Title -match "Windows 10 Version Next"}
	$Script:Win10Nextcount = $Win10Next.count	
	
	If($Win10Next)
	{
		Write-ToLog "Found $Script:Win10Nextcount Windows 10 Next Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows 10 Next Updates...";$Win10Next | %{$_.Decline()}}Else{Write-ToLog "Recording Windows 10 Next Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $Win10Next | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Win10Nexthtmfile
            If(!$TrialRun){Write-ToLog "List of Windows 10 Next updates declined: $Win10Nexthtmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Windows 10 Next updates that could be declined: $Win10Nexthtmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Windows 10 Next Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Windows 10 Next Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows 10 Next Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-ServerNext{
	$Script:ServerNextcount = 0
	Write-ToLog "$script:WsusServer is starting Decline-ServerNext function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Windows Server Next updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Windows Server Next updates..."
	$ServerNext = $GrabUpdates | where-object {$_.Title -match "Windows Server Next"}
	$Script:ServerNextcount = $ServerNext.count	
	
	If($ServerNext)
	{
		Write-ToLog "Found $Script:ServerNextcount Windows Server Next Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows Server Next Updates...";$ServerNext | %{$_.Decline()}}Else{Write-ToLog "Recording Windows Server Next Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $ServerNext | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $ServerNexthtmfile
            If(!$TrialRun){Write-ToLog "List of Windows Server Next updates declined: $ServerNexthtmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Windows Server Next updates that could be declined: $ServerNexthtmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Windows Server Next Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Windows Server Next Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows Server Next Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-IE7{
	
	
	
	Write-ToLog "$script:WsusServer is starting Decline-IE7 function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all IE7 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for IE7 updates..."
	$IE7 = $GrabUpdates | where-object {$_.Title -match "Internet Explorer 7"}
	$Script:IE7count = $IE7.count	
	
	If($IE7)
	{
		Write-ToLog "Found $Script:IE7count IE7 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining IE7 Updates...";$IE7 | %{$_.Decline()}}Else{Write-ToLog "Recording IE7 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $IE7 | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $IE7htmfile
            If(!$TrialRun){Write-ToLog "List of IE7 updates declined: $IE7htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of IE7 updates that could be declined: $IE7htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline IE7 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline IE7 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No IE7 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-IE8{

	Write-ToLog "$script:WsusServer is starting Decline-IE8 function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all IE8 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for IE8 updates..."
	$IE8 = $GrabUpdates | where-object {$_.Title -match "Internet Explorer 8"}
	$Script:IE8count = $IE8.count	
	
	If($IE8)
	{
		Write-ToLog "Found $Script:IE8count IE8 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining IE8 Updates...";$IE8 | %{$_.Decline()}}Else{Write-ToLog "Recording IE8 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $IE8 | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $IE8htmfile
            If(!$TrialRun){Write-ToLog "List of IE8 updates declined: $IE8htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of IE8 updates that could be declined: $IE8htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline IE8 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline IE8 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No IE8 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-IE9{

	Write-ToLog "$script:WsusServer is starting Decline-IE9 function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all IE9 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for IE9 updates..."
	$IE9 = $GrabUpdates | where-object {$_.Title -match "Internet Explorer 9"}
	$Script:IE9count = $IE9.count	
	
	If($IE9)
	{
		Write-ToLog "Found $Script:IE9count IE9 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining IE9 Updates...";$IE9 | %{$_.Decline()}}Else{Write-ToLog "Recording IE9 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $IE9 | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $IE9htmfile
            If(!$TrialRun){Write-ToLog "List of IE9 updates declined: $IE9htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of IE9 updates that could be declined: $IE9htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline IE9 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline IE9 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No IE9 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Embedded{

	Write-ToLog "$script:WsusServer is starting Decline-Embedded function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Embedded updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Embedded updates..."
	$Embedded = $GrabUpdates | where-object {$_.Title -match "Embedded"}
	$Script:Embeddedcount = $Embedded.count	
	
	If($Embedded)
	{
		Write-ToLog "Found $Script:Embeddedcount Embedded Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Embedded Updates...";$Embedded | %{$_.Decline()}}Else{Write-ToLog "Recording Embedded Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $Embedded | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Embhtmfile
            If(!$TrialRun){Write-ToLog "List of Embedded updates declined: $Embhtmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Embedded updates that could be declined: $Embhtmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Embedded Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Embedded Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Embedded Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-ARM64{

	Write-ToLog "$script:WsusServer is starting Decline-ARM64 function..... Please wait...."

    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all ARM64 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for ARM64 updates..."
	$Arm64 = $GrabUpdates | where-object {$_.Title -match "ARM64"}
	$Script:Arm64count = $Arm64.count	
	
	If($Arm64)
	{
		Write-ToLog "Found $Script:Arm64count ARM64-Based Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining ARM64 Updates...";$ARM64 | %{$_.Decline()}}Else{Write-ToLog "Recording ARM64 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $ARM64 | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Arm64htmfile
            If(!$TrialRun){Write-ToLog "List of ARM64 updates declined: $Arm64htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of ARM64 updates that could be declined: $Arm64htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline ARM64 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline ARM64 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No ARM64 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-IE10{

	Write-ToLog "$script:WsusServer is starting Decline-IE10 function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all IE10 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for IE10 updates..."
	$IE10 = $GrabUpdates | where-object {$_.Title -match "Internet Explorer 10"}
	$Script:IE10count = $IE10.count	
	
	If($IE10)
	{
		Write-ToLog "Found $Script:IE10count IE10 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining IE10 Updates...";$IE10 | %{$_.Decline()}}Else{Write-ToLog "Recording IE10 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $IE10 | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $IE10htmfile
            If(!$TrialRun){Write-ToLog "List of IE10 updates declined: $IE10htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of IE10 updates that could be declined: $IE10htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline IE10 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline IE10 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No IE10 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Win7{

	Write-ToLog "$script:WsusServer is starting Decline-Win7 function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Windows 7 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Windows 7 Updates..."
	$Win7=$GrabUpdates | where-object {$_.Title -match "Windows 7" -and $_.Title -notmatch "Server" -and $_.Title -notmatch "Windows 10"}
	$Script:Win7count = $Win7.count	
	
	If($Win7)
	{
		Write-ToLog "Found $Script:Win7count Windows 7 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows 7 Updates...";$Win7| %{$_.Decline()}}Else{Write-ToLog "Recording Windows 7 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $Win7| Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Win7htmfile
            If(!$TrialRun){Write-ToLog "List of Windows 7 Updates declined: $Win7htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Windows 7 Updates that could be declined: $Win7htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Windows 7 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Windows 7 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows 7 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Win8{

	Write-ToLog "$script:WsusServer is starting Decline-Win8 function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Windows 8 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Windows 8 Updates..."
	$Win8=$GrabUpdates | where-object {$_.Title -match "Windows 8" -and $_.Title -notmatch "Server" -and $_.Title -notmatch "Windows 8.1"}
	$Script:Win8count = $Win8.count	
	
	If($Win8)
	{
		Write-ToLog "Found $Script:Win8count Windows 8 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows 8 Updates...";$Win8| %{$_.Decline()}}Else{Write-ToLog "Recording Windows 8 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $Win8| Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Win8htmfile
            If(!$TrialRun){Write-ToLog "List of Windows 8 Updates declined: $Win8htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Windows 8 Updates that could be declined: $Win8htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Windows 8 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Windows 8 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows 8 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Win81{

	Write-ToLog "$script:WsusServer is starting Decline-Win81 function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Windows 8.1 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Windows 8.1 Updates..."
	$Win81=$GrabUpdates | where-object {$_.Title -match "Windows 8.1" -and $_.Title -notmatch "Server"}
	$Script:Win81count = $Win81.count	
	
	If($Win81)
	{
		Write-ToLog "Found $Script:Win81count Windows 8.1 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows 8.1 Updates...";$Win81| %{$_.Decline()}}Else{Write-ToLog "Recording Windows 8.1 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $Win81| Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Win81htmfile
            If(!$TrialRun){Write-ToLog "List of Windows 8.1 Updates declined: $Win81htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Windows 8.1 Updates that could be declined: $Win81htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Windows 8.1 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Windows 8.1 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows 8.1 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Win2k3{

	Write-ToLog "$script:WsusServer is starting Decline-Win2k3 function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Win2k3updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Windows Server 2003 Updates..."
	$Win2K3 = $GrabUpdates | where-object {$_.Title -match "Windows Server 2003" -and $_.Title -notmatch "Windows XP" -and $_.Title -notmatch "Windows 7" -and $_.Title -notmatch "Server 2008"}
	$Script:Win2k3count = $Win2k3.count	
	
	If($Win2k3)
	{
		Write-ToLog "Found $Script:Win2k3count Windows Server 2003 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows Server 2003 Updates...";$Win2k3| %{$_.Decline()}}Else{Write-ToLog "Recording Windows Server 2003 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $Win2k3| Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Win2k3htmfile
            If(!$TrialRun){Write-ToLog "List of Windows Server 2003 Updates declined: $Win2k3htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Windows Server 2003 Updates that could be declined: $Win2k3htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Windows Server 2003 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Windows Server 2003 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows Server 2003 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Win2k8{

	Write-ToLog "$script:WsusServer is starting Decline-Win2k8 function..... Please wait...."

    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Win2k8updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Windows Server 2008 Updates..."
	$Win2k8=$GrabUpdates | where-object {$_.Title -match "Windows Server 2008" -and $_.Title -notmatch "Windows 7" -and $_.Title -notmatch "Windows Server 2008 R2"}
	$Script:Win2k8count = $Win2k8.count	
	
	If($Win2k8)
	{
		Write-ToLog "Found $Script:Win2k8count Windows Server 2008 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows Server 2008 Updates...";$Win2k8| %{$_.Decline()}}Else{Write-ToLog "Recording Windows Server 2008 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $Win2k8| Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Win2k8htmfile
            If(!$TrialRun){Write-ToLog "List of Windows Server 2008 Updates declined: $Win2k8htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Windows Server 2008 Updates that could be declined: $Win2k8htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Windows Server 2008 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Windows Server 2008 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows Server 2008 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Win2k8R2{

	Write-ToLog "$script:WsusServer is starting Decline-Win2k8R2 function..... Please wait...."

    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Windows Server 2008 R2 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Windows Server 2008 R2 Updates..."
	$Win2k8R2=$GrabUpdates | where-object {$_.Title -match "Windows Server 2008 R2" -and $_.Title -notmatch "Windows 7"}
	$Script:Win2k8R2count = $Win2k8R2.count	
	
	If($Win2k8R2)
	{
		Write-ToLog "Found $Script:Win2k8R2count Windows Server 2008 R2 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows Server 2008 R2 Updates...";$Win2k8R2| %{$_.Decline()}}Else{Write-ToLog "Recording Windows Server 2008 R2 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $Win2k8R2| Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Win2k8R2htmfile
            If(!$TrialRun){Write-ToLog "List of Windows Server 2008 R2 Updates declined: $Win2k8R2htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Windows Server 2008 R2 Updates that could be declined: $Win2k8R2htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Windows Server 2008 R2 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Windows Server 2008 R2 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows Server 2008 R2 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Win12{

	Write-ToLog "$script:WsusServer is starting Decline-Win12 function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Windows Server 12 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Windows Server 12 Updates..."
	$Win12= $GrabUpdates | where-object {$_.Title -match "Windows Server 2012" -and $_.Title -notmatch "Windows Server 2012 R2" -and $_.Title -notmatch "Windows 8"}
	$Script:Win12count = $Win12.count	
	
	If($Win12)
	{
		Write-ToLog "Found $Script:Win12count Windows Server 12 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows Server 12 Updates...";$Win12| %{$_.Decline()}}Else{Write-ToLog "Recording Windows Server 12 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $Win12| Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Win12htmfile
            If(!$TrialRun){Write-ToLog "List of Windows Server 12 Updates declined: $Win12htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Windows Server 12 Updates that could be declined: $Win12htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Windows Server 12 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Windows Server 12 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows Server 12 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Win12R2{

	Write-ToLog "$script:WsusServer is starting Decline-Win12R2function..... Please wait...."


    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Windows Server 12 R2 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Windows Server 12 R2 Updates..."
	$Win12R2= $GrabUpdates | where-object {$_.Title -match "Windows Server 2012 R2" -and $_.Title -notmatch "Windows 8"}
	$Script:Win12R2count = $Win12R2.count	
	
	If($Win12R2)
	{
		Write-ToLog "Found $Script:Win12R2count Windows Server 12 R2 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows Server 12 R2 Updates...";$Win12R2| %{$_.Decline()}}Else{Write-ToLog "Recording Windows Server 12 R2 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
            $Win12R2| Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Win12R2htmfile
            If(!$TrialRun){Write-ToLog "List of Windows Server 12 R2 Updates declined: $Win12R2htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Windows Server 12 R2 Updates that could be declined: $Win12R2htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Windows Server 12 R2 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Windows Server 12 R2 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows Server 12 R2 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-LegacyWin10{

	Write-ToLog "$script:WsusServer is starting Decline-LegacyWin10 function..... Please wait...."

    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Legacy Windows 10 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Legacy Windows 10 Updates..."
	$LegacyWin10=$GrabUpdates | where-object {($_.Title -match "Windows 10 Version 1507" -and $_.Title -notmatch "Server") -or
                                                ($_.Title -match "Windows 10 Version 1511" -and $_.Title -notmatch "Server") -or 
                                                ($_.Title -match "Windows 10 Version 1607" -and $_.Title -notmatch "Server") -or 
                                                ($_.Title -match "Windows 10 Version 1703" -and $_.Title -notmatch "Server") -or 
                                                ($_.Title -match "Windows 10 Version 1803" -and $_.Title -notmatch "Server") -or 
                                                ($_.Title -match "Windows 10 Version 1903" -and $_.Title -notmatch "Server") -or 
                                                ($_.Title -match "Windows 10 Version 2004" -and $_.Title -notmatch "Server") 
                                                }
	$Script:LegacyWin10count = $LegacyWin10.count	
	
	If($LegacyWin10)
	{
		Write-ToLog "Found $Script:LegacyWin10count Legacy Windows 10 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows 10 2004 Updates...";$LegacyWin10 | %{$_.Decline()}}Else{Write-ToLog "Recording Legacy Windows 10 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
		    
            $LegacyWin10 | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $LegacyWin10htmfile
            If(!$TrialRun){Write-ToLog "List of Legacy Windows 10 Updates declined: $LegacyWin10htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Legacy Windows 10 Updates that could be declined: $LegacyWin10htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Legacy Windows 10 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Legacy Windows 10 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Legacy Windows 10 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}


Function Decline-LegacyOff365{

	Write-ToLog "$script:WsusServer is starting Decline-LegacyOff365 function..... Please wait...."

    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Legacy Office & M 365 updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Legacy Office & M 365 Updates..."
    $LegacyOff365=$GrabUpdates | where-object {($_.Title -match "Office 365" -and $_.Title -match "x86 based Edition") -or 
                                                ($_.Title -match "Microsoft 365 Apps Update" -and $_.Title -match "x86 based Edition")
                                                }
	$Script:LegacyOff365count = $LegacyOff365.count	
	
	If($LegacyOff365)
	{
		Write-ToLog "Found $Script:LegacyOff365count Legacy Office & M 365 Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Legacy Office & M 365 Updates...";$LegacyOff365 | %{$_.Decline()}}Else{Write-ToLog "Recording Legacy Office & M 365 Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
		    
            $LegacyOff365 | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $LegacyOff365htmfile
            If(!$TrialRun){Write-ToLog "List of Legacy Office & M 365 Updates declined: $LegacyOff365htmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Legacy Office & M 365 Updates that could be declined: $LegacyOff365htmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Legacy Office & M 365 Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline Legacy Office & M 365 Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Legacy Office & M 365 Updates found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-Win10FeatureUpdates{

	Write-ToLog "$script:WsusServer is starting Decline-Win10FeatureUpdates function..... Please wait...."

    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all Windows 10 Feature Updates for Enablement Package..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for Windows 10 Feature Updates for Enablement Package...."

    $FeaureWin10Updates=$GrabUpdates | where-object {($_.UpdateClassificationTitle -eq "Upgrades") -and !(($_.Title -match "x64-based") -and ($_.Title -match "Enablement Package"))}

	$Script:FeaureWin10UpdatesCount = $FeaureWin10Updates.count	
	
	If($FeaureWin10Updates)
	{
		Write-ToLog "Found $Script:FeaureWin10UpdatesCount Windows 10 Feature Updates for Enablement Package...."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining Windows 10 Feature Updates for Enablement Package...";$FeaureWin10Updates | %{$_.Decline()}}Else{Write-ToLog "Recording Windows 10 Feature Updates for Enablement Package..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}
		    
            $FeaureWin10Updates | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $FeaureWin10Updateshtmfile
            If(!$TrialRun){Write-ToLog "List of Windows 10 Feature Updates for Enablement Package declined: $FeaureWin10Updateshtmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Windows 10 Feature Updates for Enablement Package that could be declined: $FeaureWin10Updateshtmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline Windows 10 Feature Updates for Enablement Package. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Windows 10 Feature Updates for Enablement Package. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows 10 Feature Updates for Enablement Package found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}

Function Decline-OneOff{

	Write-ToLog "$script:WsusServer is starting Manual Cleanup function..... Please wait...."

	$OneOffEnable = "Disabled"

    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all updates defined for Manual Cleanup..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}

    Write-ToLog "Searching for specific Updates..."
	$script:OneOff=$GrabUpdates | where-object {$_.Title -like $SinglePatch}
	$Script:OneOffCount = $script:OneOff.count	
	
	If($script:OneOff)
	{
		Write-ToLog "Found $Script:OneOffCount OneOff Updates..."
	    Try
	    {
	        If(!$TrialRun){Write-ToLog "Declining OneOff Updates...";$script:OneOff| %{$_.Decline()}}Else{Write-ToLog "Recording OneOff Updates..."}
		    $Table = @{Name="Title";Expression={[string]$_.Title}},`
			    @{Name="KB Article";Expression={[string]$_.KnowledgebaseArticles}},`
			    @{Name="Arrival Date";Expression={[string]$_.ArrivalDate}},`
                @{Name="Classification";Expression={[string]$_.UpdateClassificationTitle}},`
			    @{Name="Product Title";Expression={[string]$_.ProductTitles}},`
			    @{Name="Product Family";Expression={[string]$_.ProductFamilyTitles}}

            $script:OneOff| Select $Table | ConvertTo-HTML -head $CStyle | Out-File $OneOffhtmfile
            If(!$TrialRun){Write-ToLog "List of Oneoff Updates declined: $OneOffhtmfile"; Write-ToLog ""}
			Else{Write-ToLog "List of Oneoff Updates that could be declined: $OneOffhtmfile"; Write-ToLog ""}

	    }
	    Catch
	    {
	        Write-EventLog -LogName $Eventlog -EventID 21031 -Message "$script:WsusServer is unable to Decline OneOff Updates. Error: $error[0]" -Source $EventSource -EntryType Error
            Write-ToLog "$script:WsusServer is unable to Decline OneOff Updates. Error: $error[0]"
	    }
	}
	Else
	{
        Write-ToLog ""
        Write-ToLog "    No Windows OneOff found that needed declining at this time..."
        Write-ToLog ""
    }
	$ErrorActionPreference = $script:CurrentErrorActionPreference

}



Function UpdateListMaint{

	Write-ToLog "CleanUpdatelist is set to $CleanUpdatelist.  Cleaning the UpdateList folder."
    Write-ToLog "Checking for logs older than $CleanULNumber day(s)."
    $oldfiles = get-childitem -Path $delpath -recurse | where-object {$_.lastwritetime -lt (get-date).addDays(-$CleanULNumber)}

    if($oldfiles)
    {
    	$numold = $oldfiles.count
        Write-ToLog "Found $numold logs and folders that are over $CleanULNumber days old."
        Write-ToLog "Deleting $numold old logs and folders."
	    Try{
		    get-childitem -Path $delpath -recurse | where-object {$_.lastwritetime -lt (get-date).addDays(-$CleanULNumber)} | Foreach-Object { remove-item $_.FullName -force -recurse}
		    get-ChildItem $delpath -recurse | Where-Object {$_.PSIsContainer -eq $True} | Where-Object {$_.GetFiles().Count -eq 0} | Foreach-Object { remove-item $_.FullName -recurse}
		    Write-ToLog "Successfully removed old logs and folders."
	    }
	    Catch
	    {
		    Write-ToLog "Failed to clean the UpdateList files/folder. Error: $Error[0]"
            Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Failed to clean the UpdateList files/folder. Error: $Error[0]" -Source $EventSource -EntryType Error
	        Write-ToLog "Please check the files manually."
		    If($EmailReport)
			    {	$Body = ConvertTo-Html -head $CStyle -Body "Failed to clean the UpdateList files/folder. Error: $Error[0].  Please check the files manually." | Out-String
				    $Body = $Body.Replace("<table>`r`n</table>", "")
				    SendEmailStatus -From $From -To $To -Subject $Subject -SmtpServer $SmtpServer -BodyAsHtml $True -Body $Body
				    Write-ToLog "Sending Mail..."
			    }

	    }

    
    }else{
    
        Write-ToLog "No old files found. Nothing to clean..."
    
    }


}



Function WSUSCleanup-CompressUpdates{
    
    Write-ToLog "Starting to run: Invoke-WsusServerCleanup -CompressUpdates Function."
    Write-ToLog "TrialRun is set to $TrialRun."
    if ($TrialRun){

        if($CompressUpdates)
        {
            Write-ToLog "This is only trialrun. No Updates would be compressed. "
            $script:Compupdates = "TrialRun"
        }else{$script:Compupdates = "Skipped"}

    
    }else{

        if($OneOffCleanup){
            Write-ToLog "OneOffCleanup is set to $OneOffCleanup. Not running CleanupCompressUpdates Function."
            $script:Compupdates = "Skipped"
            Write-ToLog "--"
        }
        else{
            if($CompressUpdates)
            {
                Write-ToLog "Running Invoke-WsusServerCleanup -CompressUpdates... "
                try
                {

                    if($UseSSL){$script:Compupdates = Get-WsusServer -name $script:WSUSServer -UseSSL -portnumber $PortNumber | Invoke-WsusServerCleanup -CompressUpdates}Else
                    {$script:Compupdates = Get-WsusServer -name $script:WSUSServer -portnumber $PortNumber | Invoke-WsusServerCleanup -CompressUpdates}

                    Write-tolog "This function is currently unavaliable."
                    $script:Compupdates = "Unavailable"

                }
                Catch
                {Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run CompressUpdates $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}

                Write-ToLog "Invoke-WsusServerCleanup -CompressUpdates Result: $script:Compupdates"
                Write-ToLog "--"
            }else
            {
                $script:Compupdates = "Skipped"
                Write-ToLog "Invoke-WsusServerCleanup -CompressUpdates Function is $script:Compupdates. Not running it."
                Write-ToLog "--"
            }
        }
    }

}

Function WSUSCleanup-CleanupObsComputers{
    
    Write-ToLog "Starting to run: Invoke-WsusServerCleanup -CleanupObsoleteComputers Function."
    Write-ToLog "TrialRun is set to $TrialRun."
    if ($TrialRun){

        if($CleanupObsoleteComputers){
            if($UseSSL){$script:CleanObsComp = (Get-WsusServer -name $script:WSUSServer -UseSSL -portnumber $PortNumber | Get-WsusComputer -ToLastSyncTime (get-date).addDays(-30)).count}
            else{$script:CleanObsComp = (Get-WsusServer -name $script:WSUSServer -portnumber $PortNumber | Get-WsusComputer -ToLastSyncTime (get-date).addDays(-30)).count}
            Write-ToLog "This is only trialrun. Otherwise number of machines would be cleaned up: $script:CleanObsComp "
        }else{$script:CleanObsComp = "Skipped"}
    
    }else{

        if($OneOffCleanup){
            Write-ToLog "OneOffCleanup is set to $OneOffCleanup. Not running CleanupObsoleteComputers Function."
            $script:CleanObsComp = "Skipped"
            Write-ToLog "--"
        }
        else{
            if($CleanupObsoleteComputers){
                Write-ToLog "Running Invoke-WsusServerCleanup -CleanupObsoleteComputers..."
                Try
                {
                    if($UseSSL){$CleanObsComptmp = Get-WsusServer -name $script:WSUSServer -UseSSL -portnumber $PortNumber | Invoke-WsusServerCleanup -CleanupObsoleteComputers}Else
                    {$CleanObsComptmp = Get-WsusServer -name $script:WSUSServer -portnumber $PortNumber | Invoke-WsusServerCleanup -CleanupObsoleteComputers}
                }
                Catch
                {Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run CleanupObsoleteComputers $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}

                $script:CleanObsComp = $CleanObsComptmp.Trim("Obsolete Computers Deleted:")

                Write-ToLog "Invoke-WsusServerCleanup -CleanupObsoleteComputers Result: $script:CleanObsComp"
                Write-ToLog "--"

            }else
            {
                $script:CleanObsComp = "Skipped"
                Write-ToLog "Invoke-WsusServerCleanup -CleanupObsoleteComputers Function is $script:CleanObsComp. Not running it."
                Write-ToLog "--"
            }
        }

    }

}

Function WSUSCleanup-UneededContentFiles{
    Write-ToLog "Starting to run: Invoke-WsusServerCleanup -CleanupUnneededContentFiles Function."
    Write-ToLog "TrialRun is set to $TrialRun."

    if ($TrialRun){

        Write-ToLog "This is only trialrun. No content files would be cleaned up. "
        if($CleanupUnneededContentFiles)
        {
            Write-ToLog ""
            Write-ToLog "This is only trialrun. No content files would be cleaned up. "
        }else{$script:CleanContents = "Skipped"}
    }else{

        if($OneOffCleanup){
            Write-ToLog "OneOffCleanup is set to $OneOffCleanup. Not running CleanupUnneededContentFiles Function."
            $script:CleanContents = "Skipped"
            Write-ToLog "--"
        }
        Else{

            if($CleanupUnneededContentFiles)
            {
                Write-ToLog "Running Invoke-WsusServerCleanup -CleanupUnneededContentFiles..."
                Try
                {

                    if($UseSSL){$CleanContents = Get-WsusServer -name $script:WSUSServer -UseSSL -portnumber $PortNumber | Invoke-WsusServerCleanup -CleanupUnneededContentFiles}Else
                    {$script:CleanContents = Get-WsusServer -name $script:WSUSServer -portnumber $PortNumber | Invoke-WsusServerCleanup -CleanupUnneededContentFiles}               

                }
                Catch
                {Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run CleanupUnneededContentFiles $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}

                Write-ToLog "Invoke-WsusServerCleanup -CleanupUnneededContentFiles Result: $script:CleanContents"
                Write-ToLog "--"

            }else
            {
                $script:CleanContents = "Skipped"
                Write-ToLog "Invoke-WsusServerCleanup -CleanupUnneededContentFile Function is $script:CleanContents. Not running it."
                Write-ToLog "--"
            }
        }

    }

    
}


############### 					###############
############### 	Main Script			###############
###############						###############


Write-ToLog "####### Starting: $scriptName #######"
Write-ToLog "####### Version: $ScriptVersion"
Write-ToLog "..."
if ($TrialRun -and $DeclineLastLevelOnly) {
    Write-ToLog "Using TrialRun and DeclineLastLevelOnly switches together is not allowed."
	Write-ToLog ""
    Write-ToLog "Exiting..."
    Exit
}

If (!(Test-Path -path $ulpath)){ $null = New-Item -ItemType Directory $ulpath -force  }


If ($MetaDataReportingOn)
{	
    Try{
		GetMetaDataSizeBefore{}
	}
	Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Error Running GetMetaDataSizeBefore Function." -Source $EventSource -EntryType Error; SendMail{}}
}
Else
{ 
	Write-ToLog "MetaData Reporting is set to $MetaDataReportingOn.  Skipping this function."

}


$jeffobjects = Foreach ($script:WsusServer in $servers)
{
	
    $error.clear()
    $StartScript = Get-Date

    $outSupersededList = "$ulpath\" + "AllSupersededUpdates-$script:WsusServer.csv"
    $outSupersededListBackup = "$ulpath\" + "SupersededUpdatesBackup-$script:WsusServer.csv"
	$outSupersededExList = "$ulpath\" + "SupersededUpdates-Over-$ExclusionPeriod-$script:WsusServer.csv"
    $outSupersededExHTM = "$ulpath\" + "SupersededUpdates-Over-$ExclusionPeriod-$script:WsusServer.htm"
    $outSupersededHTM = "$ulpath\" + "AllSupersededUpdates-$script:WsusServer.htm"
    $IThtmfile = "$ulpath\" + "Itanium-Updates-$script:WsusServer.htm"
    $XPhtmfile = "$ulpath\" + "XP-Updates-$script:WsusServer.htm"
	$Prevhtmfile = "$ulpath\" + "Preview-$script:WsusServer.htm"
	$Betahtmfile = "$ulpath\" + "Beta-$script:WsusServer.htm"
	$Win10Nexthtmfile = "$ulpath\" + "Win10Next-$script:WsusServer.htm"
	$ServerNexthtmfile = "$ulpath\" + "ServerNext-$script:WsusServer.htm"
	$IE7htmfile = "$ulpath\" + "IE7-$script:WsusServer.htm"
	$IE8htmfile = "$ulpath\" + "IE8-$script:WsusServer.htm"
	$IE9htmfile = "$ulpath\" + "IE9-$script:WsusServer.htm"
	$IE10htmfile = "$ulpath\" + "IE10-$script:WsusServer.htm"
	$Embhtmfile = "$ulpath\" + "Embedded-$script:WsusServer.htm"
    $Arm64htmfile = "$ulpath\" + "Arm64-$script:WsusServer.htm"
	$Win7htmfile = "$ulpath\" + "Win7-$script:WsusServer.htm"
	$Win8htmfile = "$ulpath\" + "Win8-$script:WsusServer.htm"
	$Win81htmfile = "$ulpath\" + "Win81-$script:WsusServer.htm"
	$Win2k3htmfile = "$ulpath\" + "Win2k3-$script:WsusServer.htm"
	$Win2k8htmfile = "$ulpath\" + "Win2k8-$script:WsusServer.htm"
	$Win2k8R2htmfile = "$ulpath\" + "Win2k8R2-$script:WsusServer.htm"
	$Win12htmfile = "$ulpath\" + "Win12-$script:WsusServer.htm"
	$Win12R2htmfile = "$ulpath\" + "Win12R2-$script:WsusServer.htm"
    $LegacyWin10htmfile = "$ulpath\" + "LegacyWin10-$script:WsusServer.htm"
	$LegacyOff365htmfile = "$ulpath\" + "LegacyOff365-$script:WsusServer.htm"
    $FeaureWin10Updateshtmfile = "$ulpath\" + "Win10Feat-$script:WsusServer.htm"
	$OneOffhtmfile = "$ulpath\" + "OneOff-$script:WsusServer.htm"

    
	Function SendMail
	{
			If($EmailReport)
				{	$Body = ConvertTo-Html -head $CStyle -Body "Error on $script:WsusServer. Error: $Error[0]" | Out-String
					SendEmailStatus -From $From -To $To -Subject $Subject -SmtpServer $SmtpServer -BodyAsHtml $True -Body $Body
					Write-ToLog "Sending Mail..."
				}
	}
	
    $Props = [ordered]@{}
	
	try {
	    
	    if ($UseSSL) {
	        Write-ToLog "Connecting to WSUS server $script:WsusServer on Port $PortNumber using SSL... "
	    } Else {
	        Write-ToLog "Connecting to WSUS server $script:WsusServer on Port $PortNumber... "
	    }
	    
		######## Connect to WSUS API########

	    [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
	    $WsusServerAdminProxy = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($script:WsusServer,$UseSSL,$PortNumber);
		

	}
	catch [System.Exception] 
	{
	    $err1 = $_.Exception.Message
		Write-ToLog "Failed to connect $script:WsusServer. Error: $err1"
        Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Error running the $scriptName Script on $script:WsusServer.  Error: $err1" -Source $EventSource -EntryType Error
	    Write-ToLog "Please check the logs."
		If($EmailReport)
			{	$Body = ConvertTo-Html -head $CStyle -Body "Failed to connect $script:WsusServer. Error: $err1.  Script exited." | Out-String
				$Body = $Body.Replace("<table>`r`n</table>", "")
				SendEmailStatus -From $From -To $To -Subject $Subject -SmtpServer $SmtpServer -BodyAsHtml $True -Body $Body
				Write-ToLog "Sending Mail..."
			}
		Write-ToLog "Exiting..."
		Exit
	}
    
    if ($WsusServerAdminProxy) {Write-ToLog "Connected to $script:WsusServer."}
    if ($TrialRun) {Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Recording all superseded updates..."}Else{Write-ToLog "NOTE: TrialRun flag is set to $TrialRun. Continuing with declining updates..."}


    $Props."Servername" = ("$script:WsusServer")

    Write-ToLog "Collecting a list of updates from $script:WsusServer... Please wait..."

    Try{$allUpdates = $WsusServerAdminProxy.GetUpdates()}
	Catch [System.Exception]
	{
	    $err1 = $_.Exception.Message
		Write-ToLog "Failed to collect a list of updates from $script:WsusServer. Error: $Error[0]"
        Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Error running the $scriptName Script on $script:WsusServer.  Error: $Error[0]" -Source $EventSource -EntryType Error
	    Write-ToLog "Please check the logs."
		If($EmailReport)
			{	$Body = ConvertTo-Html -head $CStyle -Body "Failed to collect a list of updates from $script:WsusServer. Error: $Error[0].  Please check the server manually.  Script exited." | Out-String
				$Body = $Body.Replace("<table>`r`n</table>", "")
				SendEmailStatus -From $From -To $To -Subject $Subject -SmtpServer $SmtpServer -BodyAsHtml $True -Body $Body
				Write-ToLog "Sending Mail..."
			}
		Write-ToLog "Exiting..."
		Exit
	}

	Write-ToLog "Done"

	Write-ToLog "Parsing the list of updates... " -NoNewLine

	Try{Decline-Superseded{};$Props."All Superseded" = ("$script:countSupersededAll");$Props."Superseded > $ExclusionPeriod" = ("$script:countSupersededExclusionPeriod")}
	Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-SupersededUpdatesWithExclusionPeriod function on $script:WsusServer" -Source $EventSource -EntryType Error;	SendMail{}}
		
    Write-ToLog "Decline-SupersededUpdates function.....  Done."
	Write-ToLog ""
	
	Write-ToLog "Grabbing more updates info from $script:WsusServer..."
    $GrabUpdates = $allUpdates | where-object {-not $_.IsDeclined}
	
	
	If ($OneOffCleanup)
	{	
		Try{Decline-OneOff{};$Props."OneOff" = ("$Script:OneOffcount")}
		Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-OneOff function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	}
	Else
	{ 
		Write-ToLog "Manual Cleanup is set to $OneOffCleanup.  Skipping Manual Cleanup Function."
		$1off = "Skipped"
		$Props."OneOff" = ([string]$1off)
      
		    If (!$SkipItanium)
	    {
		    Try	{
                    Decline-Itanium{}
                    if($Script:Itancount -ne '0') { $Props."Itanium" = ("$Script:Itancount") } else { $script:zerovals.Add('Itanium') > $null }
                }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-WsusItaniumUpdates function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
        Else
            { 
                Write-ToLog "SkipItanium is set to $SkipItanium.  Skipping Decline-WSusItaniumUpdates Function."
                $its = "Skipped"
                $Props."Itanium" = ([string]$its)
            }

	    If (!$SkipXP)
	    {	
		    Try{Decline-XPUpdates{}}#;$Props."Windows XP" = ($Script:XPcount)}
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-XPUpdates function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    {
		    Write-ToLog "SkipXP is set to $SkipXP.  Skipping Decline-XPUpdates Function."
		    $xps = "Skipped"
		    #$Props."Windows XP" = ([string]$xps)
	    }

	    If (!$SkipWin7)
	    {	
		    Try{
                    Decline-Win7{}
                    if($Script:Win7count -ne '0') { $Props."Win 7" = ($Script:Win7count) } else { $script:zerovals.Add('Win 7') > $null }
               }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Win7 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipWin7 is set to $SkipWin7.  Skipping Decline-Win7 Function."
		    $wn7= "Skipped"
		    {$Props."Win 7" = ([string]$wn7)}
	    }

	    If (!$SkipWin8)
	    {	
		    Try{
			    Decline-Win8{};if($Script:Win8count -ne '0'){ $Props."Win 8" = ($Script:Win8count) } else { $script:zerovals.Add('Win 8') > $null}
			    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Win8 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipWin8 is set to $SkipWin8.  Skipping Decline-Win8 Function."
		    $wn8 = "Skipped"
		    $Props."Win 8" = ([string]$wn8)
	    }

	    If (!$SkipWin81)
	    {	
		    Try{
			    Decline-Win81{};if($Script:Win81count -ne '0'){$Props."Win 8.1" = ($Script:Win81count)} else { $script:zerovals.Add('Win 8.1') > $null}
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Win81 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipWin81 is set to $SkipWin81.  Skipping Decline-Win81 Function."
		    $w81 = "Skipped"
		    $Props."Win 8.1" = ([string]$w81)
	    }

	    If (!$SkipEmbedded)
	    {	
		    Try{
			    Decline-Embedded{};if($Script:Embeddedcount -ne '0'){$Props."Embedded" = ($Script:Embeddedcount)} else { $script:zerovals.Add('Embedded') > $null} 
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Embedded function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipEmbedded is set to $SkipEmbedded.  Skipping Decline-Embedded Function."
		    $ems = "Skipped"
		    $Props."Embedded" = ([string]$ems)
	    }

	    If (!$SkipWin2k3)
	    {	
		    Try{
			    Decline-Win2k3{};if($Script:Win2k3count -ne '0'){$Props."Server 03" = ($Script:Win2k3count)} else { $script:zerovals.Add('Win 2k3') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Win2k3 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipWin2k3 is set to $SkipWin2k3.  Skipping Decline-Win2k3 Function."
		    $wk3 = "Skipped"
		    $Props."Server 03" = ([string]$wk3)
	    }

	    If (!$SkipWin2k8)
	    {	
		    Try{
			    Decline-Win2k8{};if($Script:Win2k8count -ne '0'){$Props."Server 08" = ($Script:Win2k8count)} else { $script:zerovals.Add('Win 2k8') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Win2k8 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipWin2k8 is set to $SkipWin2k8.  Skipping Decline-Win2k8 Function."
		    $wk8 = "Skipped"
		    $Props."Server 08" = ([string]$wk8)
	    }

	    If (!$SkipWin2k8R2)
	    {	
		    Try{
			    Decline-Win2k8R2{};if($Script:Win2k8R2count -ne '0'){$Props."Server 08 R2" = ($Script:Win2k8R2count)} else { $script:zerovals.Add('Win 2k8R2') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Win2k8R2 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipWin2k8R2 is set to $SkipWin2k8R2.  Skipping Decline-Win2k8R2 Function."
		    $wk8r2 = "Skipped"
		    $Props."Server 08 R2" = ([string]$wk8r2)
	    }

	    If (!$SkipWin12)
	    {	
		    Try{
			    Decline-Win12{};if($Script:Win12count -ne '0'){$Props."Server 12" = ($Script:Win12count)} else { $script:zerovals.Add('Win 12') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Win12 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipWin12 is set to $SkipWin12.  Skipping Decline-Win12 Function."
		    $w12 = "Skipped"
		    $Props."Server 12" = ([string]$w12)
	    }

	    If (!$SkipWin12R2)
	    {	
		    Try{
			    Decline-Win12R2{};if($Script:Win12R2count -ne '0'){$Props."Server 12R2" = ($Script:Win12R2count)} else { $script:zerovals.Add('Win 12R2') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Win12R2 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipWin12R2 is set to $SkipWin12R2.  Skipping Decline-Win12R2 Function."
		    $w12r2 = "Skipped"
		    $Props."Server 12R2" = ([string]$w12r2)
	    }


	    If (!$SkipPrev)
	    {	
		    Try{
			    Decline-Preview{};if($script:Prevcount -ne '0'){$Props."Previews" = ("$script:Prevcount")} else { $script:zerovals.Add('Previews') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Preview function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipPrev is set to $SkipPrev.  Skipping Decline-Preview Function."
		    $pvs = "Skipped"
		    $Props."Previews" = ([string]$pvs)
	    }
	
	    If (!$SkipBeta)
	    {	
		    Try{
			    Decline-Beta{};if($Script:Betacount -ne '0'){$Props."Beta" = ("$Script:Betacount")} else { $script:zerovals.Add('Beta') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Beta function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipBeta is set to $SkipBeta.  Skipping Decline-Beta Function."
		    $bts = "Skipped"
		    $Props."Beta" = ([string]$bts)
	    }
	

	    If (!$SkipServerNext)
	    {	
		    Try{
			    Decline-ServerNext{};if($script:ServerNextcount -ne '0'){$Props."Server Next" = ($script:ServerNextcount)} else { $script:zerovals.Add('Server Next') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-ServerNext function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipServerNext is set to $SkipServerNext.  Skipping Decline-ServerNext Function."
		    $sns = "Skipped"
		    $Props."Server 10 Next" = ([string]$sns)
	    }
	
	    If (!$SkipIE7)
	    {	
		    Try{
			    Decline-IE7{};if($Script:IE7count -ne '0'){$Props."IE 7" = ($Script:IE7count)} else { $script:zerovals.Add('IE7') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-IE7 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
        Else
            { 
                Write-ToLog "SkipIE7 is set to $SkipIE7.  Skipping Decline-IE7 Function."
                $i7s = "Skipped"
                $Props."IE 7" = ([string]$i7s)
            }
	
	    If (!$SkipIE8)
	    {	
		    Try{
			    Decline-IE8{};if($Script:IE8count -ne '0'){$Props."IE 8" = ($Script:IE8count)} else { $script:zerovals.Add('IE8') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-IE8 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
        Else
            { 
                Write-ToLog "SkipIE8 is set to $SkipIE8.  Skipping Decline-IE8 Function."
                $i8s = "Skipped"
                $Props."IE 8" = ([string]$i8s)
             }
	
	    If (!$SkipIE9)
	    {	
		    Try{
			    Decline-IE9{};if($Script:IE9count -ne '0'){$Props."IE 9" = ($Script:IE9count)} else { $script:zerovals.Add('IE9') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-IE9 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipIE9 is set to $SkipIE9.  Skipping Decline-IE9 Function."
		    $i9s = "Skipped"
		    $Props."IE 9" = ([string]$i9s)
		
	    }

	    If (!$SkipIE10)
	    {	
		    Try{
			    Decline-IE10{};if($Script:IE10count -ne '0'){$Props."IE 10" = ($Script:IE10count)} else { $script:zerovals.Add('IE10') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-IE10 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipIE10 is set to $SkipIE10.  Skipping Decline-IE10 Function."
		    $i0s = "Skipped"
		    $Props."IE 10" = ([string]$i0s)
		
	    }

	    If (!$SkipArm64)
	    {	
		    Try{
			    Decline-ARM64{};if($Script:Arm64count -ne '0'){$Props."ARM64" = ($Script:Arm64count)} else { $script:zerovals.Add('ARM64') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-ARM64 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipArm64 is set to $SkipArm64.  Skipping Decline-ARM64 Function."
		    $arm = "Skipped"
		    $Props."ARM64" = ([string]$arm)
	    }

	    If (!$SkipWwin10Next)
	    {	
		    Try{
			    Decline-Win10Next{};if($Script:Win10Nextcount -ne '0'){$Props."Win10 Next" = ("$Script:Win10Nextcount")} else { $script:zerovals.Add('Win10 Next') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Win10Next function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipWin10Next is set to $SkipWin10Next.  Skipping Decline-Win10Next Function."
		    $wns = "Skipped"
		    $Props."Win10 Next" = ([string]$wns)
	    }



	    If (!$SkipLegacyWin10)
	    {	
		    Try{
			    Decline-LegacyWin10{};if($Script:LegacyWin10count -ne '0'){$Props."Legacy Win10" = ("$Script:LegacyWin10count")} else { $script:zerovals.Add('Legacy Win10') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-LegacyWin10 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipLegacyWin10 is set to $SkipLegacyWin10.  Skipping Decline-LegacyWin10 Function."
		    $LGWin10 = "Skipped"
		    $Props."Legacy Win10" = ([string]$LGWin10)
        }
    

        
        If (!$SkipLegacyOff365)
	    {	
		    Try{
			    Decline-LegacyOff365{};if($Script:LegacyOff365count -ne '0'){$Props."Legacy O365" = ("$Script:LegacyOff365count")} else { $script:zerovals.Add('Legacy O365') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-LegacyOff365 function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipLegacyOff365 is set to $SkipLegacyOff365.  Skipping Decline-LegacyOff365 Function."
		    $LGOff365 = "Skipped"
		    $Props."Legacy O365" = ([string]$LGOff365)
        }
    
    
        If (!$SkipWin10FeatureUpdates)
	    {	
		    Try{
			    Decline-Win10FeatureUpdates{};if($Script:FeaureWin10UpdatesCount -ne '0'){$Props."Win10 Servicing Updates" = ("$Script:FeaureWin10UpdatesCount")} else { $script:zerovals.Add('Win10 Servicing Updates') > $null }
		    }
		    Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run Decline-Win10FeatureUpdates function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	    }
	    Else
	    { 
		    Write-ToLog "SkipWin10FeatureUpdates is set to $SkipWin10FeatureUpdates.  Skipping  Decline-Win10FeatureUpdates Function."
		    $Win10FUP = "Skipped"
		    $Props."Win10 Servicing Updates" = ([string]$Win10FUP)
        }

    }

    Write-ToLog "Done with All of the Decline updates functions."
    Write-ToLog "-"
    Write-ToLog "-"
    Write-ToLog "Starting the cleanup SUSDB functions on $script:WsusServer..."
	
	Try{WSUSCleanup-CompressUpdates{};$Props."Cleanup Compress Updates" = ($script:Compupdates)}
	Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run WSUSServerCleanup function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	Try{WSUSCleanup-CleanupObsComputers{};$Props."Cleanup Obsolete Computers" = ($script:CleanObsComp)}
	Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run WSUSServerCleanup function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}
	Try{WSUSCleanup-UneededContentFiles{};$Props."Clean Uneeded Content Files" = ($script:CleanContents)}
	Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Unable to run WSUSServerCleanup function on $script:WsusServer" -Source $EventSource -EntryType Error; SendMail{}}

    Write-ToLog "Done cleaning up the WSUSServer DB... "
    if ($error)
    {
        Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Error running the $scriptName Script on $script:WsusServer.  Error: $error{0}" -Source $EventSource -EntryType Error
		SendMail{}
    }
    Else
    {

     Write-EventLog -LogName $Eventlog -EventID 21020 -Message "$scriptName Script has completed successfully on $script:WsusServer." -Source $EventSource -EntryType Information

     $xcounter++

     if($scount -eq $xcounter)
     { 

            If ($MetaDataReportingOn)
            {	
                Try{
		            GetMetaDataSizeAfter{}
	            }
	            Catch{Write-EventLog -LogName $Eventlog -EventID 21021 -Message "Error Running GetMetaDataSizeBefore Function." -Source $EventSource -EntryType Error; SendMail{}}
            }
            Else
            { 
	            Write-ToLog "MetaData Reporting is set to $MetaDataReportingOn.  Skipping this function."

            }
     }

    }

    #$row = $table.NewRow()

	Write-ToLog ""
	Write-ToLog "Done with $script:WsusServer... "
	Write-ToLog ""
	Write-ToLog "=========================="
	Write-ToLog "Overall Decline Cleanup Summary for $script:WsusServer"
    Write-ToLog "    NOTE: TrialRun flag is set to $TrialRun"
	Write-ToLog "=========================="
	Write-ToLog "    Total Superseded Updates (Older than $ExclusionPeriod days) = $script:countSupersededExclusionPeriod"
	Write-ToLog "    Total Itanium Updates = $script:Itancount. $its"
	Write-ToLog "    Total Windows XP Updates = $Script:XPcount. $xts"
	Write-ToLog "    Total Windows 7 Updates = $Script:Win7count. $wn7"
	Write-ToLog "    Total Windows 8 Updates = $Script:Win8count. $wn8"
	Write-ToLog "    Total Windows 8.1 Updates = $Script:Win81count. $w81"
	Write-ToLog "    Total Windows Server 03 Updates = $Script:Win2k3count. $wk3"
	Write-ToLog "    Total Windows Server 08 Updates = $Script:Win2k8count. $wk8"
	Write-ToLog "    Total Windows Server 08 R2 Updates = $Script:Win2k8R2count. $wk8r2"
	Write-ToLog "    Total Windows Server 12 Updates = $Script:Win12count. $w12"
	Write-ToLog "    Total Windows Server 12 R2 Updates = $Script:Win12R2count. $w12r2"
	Write-ToLog "    Total Windows Embedded Updates = $Script:Embeddedcount. $ems"
	Write-ToLog "    Total Windows Server Next Updates = $Script:ServerNextcount. $sns"
	Write-ToLog "    Total Preview Updates = $Script:Prevcount. $pvs"
	Write-ToLog "    Total Beta Updates = $Script:Betacount. $bts"
	Write-ToLog "    Total ARM64-Based Updates = $Script:Arm64count. $arm"
	Write-ToLog "    Total IE 7 Updates = $Script:IE7count. $i7s"
	Write-ToLog "    Total IE 8 Updates = $Script:IE8count. $i8s"
	Write-ToLog "    Total IE 9 Updates = $Script:IE9count. $i9s"
	Write-ToLog "    Total IE 10 Updates = $Script:IE10count. $i0s"
    Write-ToLog "    Total Legacy Windows 10 Updates = $Script:LegacyWin10count. $LGWin10"
	Write-ToLog "    Total Legacy Office and M365 Updates = $Script:LegacyOff365count. $LGOff365"
    Write-ToLog "    Total Win10 Servicing Updates / Enablement Pacakges = $Script:FeaureWin10UpdatesCount. $Win10FUP"
	Write-ToLog "    Total OneOff Updates = $Script:OneOffcount. $1off"

	Write-ToLog ""
	Write-ToLog "=========================="
	Write-ToLog "WSUS Cleanup Results for $script:WsusServer"
	Write-ToLog "=========================="
	Write-ToLog "    Invoke-WsusServerCleanup -CompressUpdates Result: = $script:Compupdates"
	Write-ToLog "    Invoke-WsusServerCleanup -CleanupObsoleteComputers Result: = $script:CleanObsComp."
	Write-ToLog "    Invoke-WsusServerCleanup -CleanupUnneededContentFiles Result: = $script:CleanContents."

	Write-ToLog ""
	Write-ToLog "=========================="
	Write-ToLog "WSUS Metadata size and Update Counts Comparison: "
    if($MetaDataReportingOn){
        Write-ToLog "WSUS MetaData Reporting is set to $MetaDataReportingOn."
	    Write-ToLog " BEFORE:"
	    Write-ToLog "--- Updates Count: $script:SbeforeUpdateCount"
	    Write-ToLog "--- Catalog Size (MB): $Script:SbeforeCatSize"
	    Write-ToLog "--- Compressed Catalog Size (MB): $Script:SbeforeCompCatSize"
	    Write-ToLog " AFTER:"
	    Write-ToLog "--- Updates Count: $script:SAfterUpdateCount"
	    Write-ToLog "--- Catalog Size (MB): $script:SAfterCatSize"
	    Write-ToLog "--- Compressed Catalog Size (MB): $script:SAfterCompCatSize"
	    Write-ToLog "=========================="	
    }else{Write-ToLog "WSUS MetaData Reporting is set to $MetaDataReportingOn. Skipping this report."}

    
	Write-ToLog ""
	If (!$TrialRun){Write-ToLog "These Updates were declined, unless Skipped."}Else{Write-ToLog "Updates were ONLY recorded. See UpdatesList folder."}
	Write-ToLog "=========================="
	Write-ToLog "=========================="
	Write-ToLog ""

    $Props."TrialRun" = ($TrialRun)
   
    $StopScript = Get-Date
    $timespan = new-timespan -seconds $(($StopScript-$startScript).totalseconds) 
    $ScriptTime = '{0:00}h:{1:00}m:{2:00}s' -f $timespan.Hours,$timespan.Minutes,$timespan.Seconds
	Write-ToLog "$script:WsusServer Run Time: $ScriptTime"
	Write-ToLog "=========================="
	Write-ToLog ""

    $Props."Run Time" = ($ScriptTime)

    New-Object PSObject -property $Props
}

$jeffobjects | Select $Table | ConvertTo-HTML -head $CStyle | Out-File $Overallhtmfile

If($CleanUpdatelist){UpdateListMaint{}}


if($OneOffCleanup){
    

    $PrintOneOff = $script:OneOff.Title -join ", "
	If($EmailReport)
	{	
		$Body = "<p class=MsoTitle><span style='font-size:20pt;font-family:Verdana,sans-serif'>$ReportTitle<o:p></o:p></span></p>"
		$Body += $jeffobjects | Select $Table | ConvertTo-HTML -head $CStyle -PostContent "
			<h5>One off is set to $OneOffCleanup and below are declined
			<h6>   $PrintOneOff
			<h6>Happy Patching!
			<h6>Created $(Get-Date)</h6><br>$from"
		SendEmailStatus -From $From -To $To -Subject $Subject -SmtpServer $SmtpServer -BodyAsHtml $True -Body $Body
	}


}else{



		$script:zerovals = $script:zerovals | select -Unique
		$zeroval = $script:zerovals -join ", "
		
	
		If($EmailReport)
			{	
				if($MetaDataReportingOn){
                
                    $Body = "<p class=MsoTitle><span style='font-size:20pt;font-family:Verdana,sans-serif'>$ReportTitle<o:p></o:p></span></p>"
				    $Body += $jeffobjects | Select $Table | ConvertTo-HTML -head $CStyle -PostContent "
					    <h5>The following have 0 values
					    <h6>   $zeroval
                        <h5>WSUS MetaData Size Report
                        <h6>	BEFORE:
	                    <h6>        --- Update Count: $script:SbeforeUpdateCount
	                    <h6>        --- Catalog Size (MB): $Script:SbeforeCatSize
	                    <h6>        --- Compressed Catalog Size (MB): $Script:SbeforeCompCatSize
	                    <h6>    AFTER:
	                    <h6>        --- Update Count: $script:SAfterUpdateCount
	                    <h6>        --- Catalog Size (MB): $script:SAfterCatSize
	                    <h6>        --- Compressed Catalog Size (MB): $script:SAfterCompCatSize
					    <h5>Happy Patching!
					    <h6>Created $(Get-Date)</h6><br>$from"
				    SendEmailStatus -From $From -To $To -Subject $Subject -SmtpServer $SmtpServer -BodyAsHtml $True -Body $Body
                }
                Else
                {
                                    $Body = "<p class=MsoTitle><span style='font-size:20pt;font-family:Verdana,sans-serif'>$ReportTitle<o:p></o:p></span></p>"
				    $Body += $jeffobjects | Select $Table | ConvertTo-HTML -head $CStyle -PostContent "
					    <h5>The following have 0 values
					    <h6>   $zeroval
					    <h5>Happy Patching!
					    <h6>Created $(Get-Date)</h6><br>$from"
				    SendEmailStatus -From $From -To $To -Subject $Subject -SmtpServer $SmtpServer -BodyAsHtml $True -Body $Body
                }
			}

}

  
if($scount -eq $xcounter)
{

    if($error)
    {
          Write-ToLog "There are errors. Not forcing a SUP sync. Please check the logs."
          Write-ToLog "Error: $error[0]."
    }
    else
    {
        if(!$Trialrun)
        {
            if($forcesync)
                {
                    Write-ToLog "Force Synch is set to $forcesync. Forcing a SUP sync."
                    $SUP = [wmiclass]("\\$CMProvider\root\SMS\Site_$($SiteCode):SMS_SoftwareUpdate")
                    $Params = $SUP.GetMethodParameters("SyncNow")
                    $Params.fullSync = $true
                    $Return = $SUP.SyncNow($Params)
        
                }
            Else
            {
                Write-ToLog "Force Synch is set to $forcesync. Not forcing a SUP sync."
            }
        }
        else
        {
            Write-ToLog "Trialrun is set to $Trialrun. Not forcing a SUP sync."
        }  

    }
	

}


Write-ToLog "====>  Done  <===="
Write-ToLog "Script: $scriptName #######"
Write-ToLog "Ver: $ScriptVersion"
Write-ToLog "All target WSUS/SUP servers have been completed."
Write-ToLog "See $Overallhtmfile."
Write-ToLog "=========================================="



