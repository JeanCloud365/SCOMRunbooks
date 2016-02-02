import-module operationsmanager -global -force

function Load-cSCOMSDK {
	<#
    .SYNOPSIS 
      Loads the SCOM libraries from the GAC. A SCOM console must be installed for this to work!
    
  #>

	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.OperationsManager.Common") | Out-Null

	[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.OperationsManager") | Out-Null


}
Load-cSCOMSDK

function Remove-cSCOMObject {
	<#
    .SYNOPSIS 
      Removes a monitoring object from the management group
    .DESCRIPTION
     This function uses the official Operations Manager 2012 SDK to remove monitoring objects (which can be piped to it). This is a very powerful tool which can
	cause a lot of problems if used incorrectly! You must see this as a (supported) last resort, and try the Remove-SCOMDisabledClassInstance first.
  #>
	param(
	# The monitoringobject to remove
	[Parameter(ValueFromPipeline=$true)]
	[Microsoft.EnterpriseManagement.Monitoring.MonitoringObject] $Object,
	# When using this switch, there won't be a prompt before deletion
	[switch] $Force
	)
	process{
		foreach($o in $Object){
			$idd = new-object Microsoft.EnterpriseManagement.ConnectorFramework.IncrementalDiscoveryData
			$idd.remove($o)
			if($Force -eq $false){
				if($(Read-Host "This will remove the object $($o.FullName) from Management Group $($o.ManagementGroup.Name). Continue? [Y/N]").ToLower() -ne 'y'){
					return
				}
			}
			$idd.commit($o.ManagementGroup)
		}
	} 
}