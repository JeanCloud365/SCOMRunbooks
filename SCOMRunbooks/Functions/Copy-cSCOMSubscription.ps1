# This function copies an entire subscription from a source management group to a target one. If the subscription or one of its referenced subscribers and channels already exist in the target MG, they are re-used, else they are created.
# The function currently only supports a single mail or command channel notification-action!
# You need to run this function with an account that has admin rights in both environments!
# Copy-cSCOMSubscription -subscription "mysubscription" -source scom-dev-server -target scom-prod-server

function Copy-cSCOMSubscription
{

param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    # the display-name of the subscription to copy
	$subscription,
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    # The hostname of a SCOM management server belonging to the source Management Group
	$source,
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    # The hostname of a SCOM management server belonging to the target Management Group
	$target
    )
	$ErrorActionPreference = "stop"

    # import the SCOM module to obtain the standard cmdlets
	if($root = Get-ItemPropertyValue -Path "HKLM:SOFTWARE\Microsoft\System Center Operations Manager\12\Setup\Console" -Name InstallDirectory -ErrorAction SilentlyContinue){
    	if(!(get-module -name operationsmanager)){
			import-module "$root\..\Powershell\OperationsManager\OperationsManager.psm1"
		}
	} else {
		throw "SCOM module not found"
	}
	
	# create empty subscription, subscriber and channel vars
	$ch = $null
	$ss = @()
	$sub = $null
	
	# get the source subscription
	$sourcesub = Get-SCOMNotificationSubscription -DisplayName "$($subscription)" -computername $source
	# check if the first defined channel used in the source subscription already exists in the target MG (based on displayname).
	# If it does, it is reused, else a new one is created using the same parameters as the source one. Only supports mail or command channels currently.
	if(($ch = Get-SCOMNotificationChannel -DisplayName $sourcesub.Actions[0].DisplayName -ErrorAction SilentlyContinue -computername $target) -ne $null){
		
	} else {
		if($sourcesub.Actions[0].GetType().Name -eq 'CommandNotificationAction'){
			$ch = Add-SCOMNotificationChannel -computername $target -Name $sourcesub.Actions[0].Name -DisplayName $sourcesub.Actions[0].DisplayName -Argument $sourcesub.Actions[0].CommandLine -ApplicationPath $sourcesub.Actions[0].ApplicationName -WorkingDirectory $sourcesub.Actions[0].WorkingDirectory
		}
		elseif($sourcesub.Actions[0].GetType().Name -eq 'SmtpNotificationAction'){
			$ch = Add-SCOMNotificationChannel -computername $target -Name $sourcesub.Actions[0].Name -DisplayName $sourcesub.Actions[0].DisplayName -Subject $sourcesub.Actions[0].Subject -Body $sourcesub.Actions[0].Body -From $sourcesub.Actions[0].From -Server $SourceSub.Actions[0].Endpoint.PrimaryServer.Address -ReplyTo $sourcesub.Actions[0].From
		}
	}
	# parse through all the subscription receipients aka subscribers. Add existing ones to the empty array $ss. If a source subscriber does not exist, it is created with the correct
	foreach($i in $sourcesub.ToRecipients){
	
		if(($j = Get-SCOMNotificationSubscriber -computername $target -Name $i.Name -ErrorAction SilentlyContinue) -ne $null){
			$ss += $j
		} 	else {
			# to simplify the subscriber creation code, we add a dummy value which we replace later by the real subscriber device (mail, command).
			
			$j = Add-SCOMNotificationSubscriber -DeviceList "dummy@dummy.dummy" -Name $i.Name -computername $target
			$j.devices.clear()
           
            $action += New-Object -TypeName Microsoft.EnterpriseManagement.Administration.NotificationRecipientDevice -ArgumentList $($i.devices[0].Protocol),$($i.devices[0].Address)
            $action.name = $i.devices[0].name
			$j.devices.add($action)
			$j.update()
			$ss += $j
		}
	}
	# finally, we bring it all together in the subscription code.
	# First we check if the subscription already exist (displayname)
	# If it does, we only copy the subscription configuration (which alerts are forwarded) without touching subscribers and channels
	# Else we create a new subscription using the channel and subscriber-array variables we populated, and fill in the gaps with values from the source subscription
	# the alert criteria is always copied over as mentioned before
	# the subscription is created in a disabled state, you should enable it manually with PS after optional validations
	if(($sub = Get-SCOMNotificationSubscription -DisplayName $sourcesub.DisplayName -computername $target -ErrorAction SilentlyContinue) -ne $null){	
		$sub.configuration = $sourcesub.configuration
		$sub.update()
	} else {
		$sub = Add-SCOMNotificationSubscription -computername $target -Name $sourcesub.Name -Channel $ch -Disabled -Subscriber $ss -DisplayName $sourcesub.DisplayName -Delay 10 -Description $sourcesub.description -Criteria $sourcesub.configuration.criteria
		$sub.configuration = $sourcesub.configuration
		$sub.update()
	}



}

Copy-cSCOMSubscription -subscription "mysubscription" -source scom-dev-server -target scom-prod-server