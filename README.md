# PS
.SYNOPSIS
Script is for declining superseeded, Itanium, Preview, Beta, ARM64, IE7, IE8, IE9, IE10, Win10 Next, Server Next, Embedded, Legacy Win10, Unwanted M365 updates, and various Legacy Windows OS (See below) Updates in WSUS/SUP environment, and MORE!!!


    It can also now target and decline individual patches, decline Win10 servicing updates, and compare metadata before and after...
	
    Recommend running this monthly...  
	Run the scripts targetting the bottom or downstreams servers (bottom SUPs), then run it against the upstream server (Top SUP)...
	BE AWARE:  
		The OS filtering function may grab Other articles. I HIGHLY recommend doing a -TrialRun first, and examine the results in the 'UpdatesList' folder before executing this in prod environment.  
			To do this, either run the script with -TrialRun or set this switch to $true (which is on by default), see param section below.  
			AND remove the = $true off the OS/products' switches below if you'd like to see list of the updates that you'd like to decline first, before setting the -TrialRun to false.

.DESCRIPTION 
 Script is designed to decline all of the updates that have been superseded for over 90 days (by default), and MORE!!!. 

Latest Updates from previous release (5.0)****
	Version:	5.4.3
	Author:		Jeff Carreon
    NEW! Updates: ver. 5.4.3  (12/17/2021)
        - Added Metadatasize comparison report (Requires Importing SQLPS or SQLServer Module)
    NEW! Updates: ver. 5.4.2  (12/15/2021)
        - Added a function to decline Windows 10 Feature Updates for Enablement package
        - Added -forcesync.  For forcing SUP synchronization from Top down on CM Hierarchies
        - Fixed the CleanUpdateList function
        
    NEW! Updates: ver. 5.4.1  (3/9/2021)
        - Updated the report to not show categories with 0 results.  Though it will list the ones with 0 below the table.
	  NEW! Updates: ver. 5.4  (3/2/2021)
		- Added an OneOff Manual Decline function.  For declining single patches or multiple depending on the -kb input (below)
        - example usage:  .\Run-DeclineUpdate-Cleanup.ps1 -trialrun -OneOffCleanup -kb "*KB2768005*"
        NOTEs: 
            - The -OneOffCleanup depends on -kb being populated
            - The -kb uses the "like" operator. 
            - I strongly recommened using the -Trialrun first, then validate the list of patches that are documented in the logs and html it creates before declining.
	  NEW! Updates: ver. 5.3  (12/9/2020)
		- Added a function for Windows 10 versions 1507/1511/1607/1703/1803/1903/2004Â Â 
        - Added a function for declining legacy Office and M365
    NEW! Updates: ver. 5.2  (10/14/2020)
        - Added the following, but only -CleanupObsoleteComputers is being used.
        	Invoke-WsusServerCleanup -CompressUpdates
	        Invoke-WsusServerCleanup -CleanupObsoleteComputers
	        Invoke-WsusServerCleanup -CleanupUnneededContentFiles

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
