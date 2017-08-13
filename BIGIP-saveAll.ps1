$pshost = Get-Host
$psWindow = $pshost.UI.RawUI
$newSize =$psWindow.BufferSize
$newSize=(get-host).UI.RawUI.MaxWindowSize
$psWindow.WindowSize= $newSize


$psWindow.WindowSize= $newSize
$ScriptPath = Split-Path $MyInvocation.InvocationName
$cmd = "$ScriptPath\BIGIP2Excel_v2.ps1"


ForEach ($content in Get-Content "$ScriptPath\ip.txt")
{
$args = @()
$args += ("-mgmt_user", "svc-backup")
$args += ("-mgmt_pass", "HgDf2ctrp!P")
$args += ("-mgmt_ip", "$content")
Invoke-Expression "$cmd $args"

}



