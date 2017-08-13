#REQUIRES -Version 3.0

<# 
.SYNOPSIS
    BIGIP2Excel_v2.ps1
    Script for inventory BIGIP LTM
.DESCRIPTION 
    Version: 1.0
    Date: 10/05/2017
    Author: Guillaume HUGUES
.PARAMETER mgmt_ip
    BIGIP address
.PARAMETER mgmt_user
    User name
.PARAMETER mgmt_pass
    User password
#>
[CmdletBinding()]
param (
	[Net.IPAddress]$mgmt_ip = $null,
	[string]$mgmt_user = $null,
	[string]$mgmt_pass = $null
)

Add-Type @"
	using System.Net;
	using System.Security.Cryptography.X509Certificates;
	public class ISkipUntrustedCert : ICertificatePolicy {
		public ISkipUntrustedCert() {}
		public bool CheckValidationResult(
			ServicePoint sPoint, X509Certificate cert, WebRequest wRequest, int certProb) {
			return true;
		}
	}
"@
[System.Net.ServicePointManager]::CertificatePolicy = new-object ISkipUntrustedCert


#region Global

Set-PSDebug -strict;

Set-Variable header_BGColor -option ReadOnly -value 2162853
Set-Variable header_ForeColor -option ReadOnly -value 1

Set-Variable colNode_Name -option ReadOnly -value "A"
Set-Variable colNode_Address -option ReadOnly -value "B"
Set-Variable colNode_Monitor -option ReadOnly -value "C"
Set-Variable colNode_ConnLimit -option ReadOnly -value "D"

Set-Variable colPool_Name -option ReadOnly -value "A"
Set-Variable colPool_Member -option ReadOnly -value "B"
Set-Variable colPool_LBMethod -option ReadOnly -value "C"
Set-Variable colPool_Monitor -option ReadOnly -value "D"

Set-Variable colVS_Name -option ReadOnly -value "A"
Set-Variable colVS_Source -option ReadOnly -value "B"
Set-Variable colVS_Destination -option ReadOnly -value "C"
Set-Variable colVS_Port -option ReadOnly -value "D"
Set-Variable colVS_Protocol -option ReadOnly -value "E"
Set-Variable colVS_Description -option ReadOnly -value "F"
Set-Variable colVS_DefaultPool -option ReadOnly -value "G"
Set-Variable colVS_Persistence -option ReadOnly -value "H"
Set-Variable colVS_FallbackPersist -option ReadOnly -value "I"
Set-Variable colVS_SNAT -option ReadOnly -value "J"
Set-Variable colVS_SNATPool -option ReadOnly -value "K"
Set-Variable colVS_VLAN -option ReadOnly -value "L"
Set-Variable colVS_ClientSSL -option ReadOnly -value "M"
Set-Variable colVS_ServerSSL -option ReadOnly -value "N"
Set-Variable colVS_TCPClient -option ReadOnly -value "O"
Set-Variable colVS_TCPServer -option ReadOnly -value "P"
Set-Variable colVS_HTTPProfile -option ReadOnly -value "Q"
Set-Variable colVS_OthProfile -option ReadOnly -value "R"
Set-Variable colVS_iRule -option ReadOnly -value "S"
Set-Variable colVS_Policy -option ReadOnly -value "T"

$script:device = $null;
#endregion

function Get-MaskLength($mask) {
    try {
        $len = "$( $mask.Split(".") | ForEach-Object { [Convert]::ToString($_, 2) } )" -replace '[\s0]';
        return [string]$len.Length;
    } catch {
        return "0";
    }
}

#region Declared functions
<#
.DESCRIPTION
	This function gets rest method
#>
function Invoke-Get($uri) {
	$response = Invoke-RestMethod -Uri "$($script:device.serviceUrl)$uri" -Credential $script:device.credential -Method Get -Verbose:($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent -eq $true);
    Write-Verbose "REST $response";
    return $response;
}

<#
.DESCRIPTION
	This function check for initialization
#>
function Initialize() {
(get-host).UI.RawUI.MaxWindowSize
Write-Host "///////////////////////////////////////////////////////////////"
Write-Host "// Exporting configuration script for F5 BigIP               //"
Write-Host "// Created by MCO Security to survey F5 Platform             //"
Write-Host "// In the frame of the contract MCO Network Security         //"
Write-Host "// Operated by company SFR Business Solutions (Cédric JULIEN)//"
Write-Host "// Email: support.network-security.ah@airbus.com             //"
Write-Host "///////////////////////////////////////////////////////////////"
Write-Host " "
Write-Host "///////////////////////////////////////////////////////////////////////////////////"
Write-Host "// If you have a problem for launch this script.                                 //"
Write-Host "// Please launch this command in Administrator: set-executionpolicy unrestricted //"
Write-Host "///////////////////////////////////////////////////////////////////////////////////"
Write-Host " "
Write-Host "////////////////////////////////////////////////////////////////////////////////"
Write-Host "// You can launch this script with this argument                              //"
Write-Host "// .\BIGIP2Excel_v2.ps1 -mgmt_user <user> -mgmt_ip <IP> -mgmt_pass <Password> //"
Write-Host "////////////////////////////////////////////////////////////////////////////////"


	if ([String]::IsNullOrEmpty($mgmt_ip)) {
		$mgmt_ip = Read-Host "`nManagement IP";
        if ($mgmt_ip -eq "") {
            exit;
        }
	}

	if ([String]::IsNullOrEmpty($mgmt_user)) {
		$mgmt_user = Read-Host "User";
        if ($mgmt_user -eq "") {
            exit;
        }
	}

	if ([String]::IsNullOrEmpty($mgmt_pass)) {
		$mgmt_pass = Read-Host -assecurestring "Password";
    } else {
        $mgmt_pass = (ConvertTo-SecureString –String $mgmt_pass –AsPlainText -Force);
    }

    $script:device = @{
	    credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $mgmt_user, $mgmt_pass;
	    serviceUrl = "https://$mgmt_ip/mgmt/tm";
        settings = $null;
    }

    Write-Host "`nRetrieving device settings...";
	$script:device.settings = Invoke-Get "/sys/global-settings/";

    if ($script:device.settings) {
        Write-Host "Device: $($script:device.settings.hostname)";
        return $true;
    }

	return $false;
}


function Create-Header($WorkSheet) {

	$row = 1;
	$WorkSheet.Cells.Item($row, 1) = "BIGIP Inventory";
	$WorkSheet.Cells.Item($row, 1).Font.Bold = $true;
	$WorkSheet.Cells.Item($row, 1).Font.Size = 18;
	$row++;

	$WorkSheet.Cells.Item($row, 1) = "Device: " + $script:device.settings.hostname;
	$WorkSheet.Cells.Item($row, 1).Font.Size = 14;
	$WorkSheet.Cells.Item($row, 2) = "Date&Time: " + (Get-Date -Format "MM/dd/yyyy HH:mm:ss");

	$row+=2;
    $WorkSheet.Cells.Item($row, "A") = $WorkSheet.Name;
	$WorkSheet.Cells.Item($row, 1).Font.Size = 16;

    $header = $WorkSheet.Rows.Item("$row`:$($row+1)").EntireRow;
    $header.Interior.Color = $header_BGColor;
    $header.Font.ThemeColor = $header_ForeColor;
    $header.Font.Bold = $true;

    return [int]++$row;
}

<#
.DESCRIPTION
	This function create and fills worksheet for nodes
#>
function Export-WorkSheetNode($WorkBook) {

	$WorkSheet = $WorkBook.WorkSheets.Item(1);
	$WorkSheet.Name = "Node";

    $row = Create-Header($WorkSheet);
    $WorkSheet.Cells.Item($row, $colNode_Name) = "Name";
    $WorkSheet.Cells.Item($row, $colNode_Address) = "Address";
    $WorkSheet.Cells.Item($row, $colNode_Monitor) = "Monitor";
    $WorkSheet.Cells.Item($row, $colNode_ConnLimit) = "Connection Limit";
    $row++;
    $x = $WorkSheet.Cells.Item("$row","B").Select();
    $Excel.ActiveWindow.FreezePanes = $true;
	
	Write-Host "Retrieving node list...";
	$node_list = Invoke-Get "/ltm/node/?`$select=fullPath,address,connectionLimit,monitor";
	
	Write-Host "Filling spreadsheet..." -NoNewline
	foreach ($node in $node_list.items) {
        Write-Verbose "Node: $node";

		$WorkSheet.Cells.Item($row, $colNode_Name) = $node.fullPath;
		$WorkSheet.Cells.Item($row, $colNode_Address) = $node.address;
		$WorkSheet.Cells.Item($row, $colNode_Monitor) = $node.monitor;
		$WorkSheet.Cells.Item($row, $colNode_ConnLimit) = $node.connectionLimit;
		
		Write-Host "." -NoNewline;
		$row++;
	}
	$x = $WorkSheet.Columns.Item("A:Z").EntireColumn.AutoFit();
	$x = $WorkSheet.Columns.Item("A:A").EntireRow.AutoFit();

	Write-Host "";
}

<#
.DESCRIPTION
	This function create and fills worksheet for pools
#>
function Export-WorkSheetPool($WorkBook) {

	$WorkSheet = $WorkBook.WorkSheets.Add();
	$WorkSheet.Name = "Pool";
    $row = Create-Header($WorkSheet);
	$WorkSheet.Cells.Item($row, $colPool_Name) = "Name";
	$WorkSheet.Cells.Item($row, $colPool_Member) = "Member";
	$WorkSheet.Cells.Item($row, $colPool_LBMethod) = "Load Balance Method";
	$WorkSheet.Cells.Item($row, $colPool_Monitor) = "Monitor";
    $row++;
    $x = $WorkSheet.Cells.Item("$row","B").Select();
    $Excel.ActiveWindow.FreezePanes = $true;
	
	Write-Host "Retrieving pool list...";
	$pool_list = Invoke-Get "/ltm/pool/?`$select=name,partition,fullPath,loadBalancingMode,monitor";

	Write-Host "Filling spreadsheet..." -NoNewline;
	foreach ($pool in $pool_list.items) {
        Write-Verbose "Pool: $pool";

		$WorkSheet.Cells.Item($row, $colPool_Name) = $pool.fullPath;
		$members = Invoke-Get "/ltm/pool/$($pool.fullPath.Replace('/', '~'))/members/?`$select=fullPath, items";
		for($i=0; $i -lt $members.items.count; $i++) {
			$member = $members.items[$i];
			$WorkSheet.Cells.Item($row, $colPool_Member) = $member.fullPath;
		}
		$WorkSheet.Cells.Item($row, $colPool_LBMethod) = $pool.loadBalancingMode;
		$WorkSheet.Cells.Item($row, $colPool_Monitor) = $pool.monitor;
		
		Write-Host "." -NoNewline;
		$row++;
	}
	$WorkSheet.Columns.Item($colPool_member).ColumnWidth = 50;
	$x = $WorkSheet.Columns.Item("A:Z").EntireColumn.AutoFit();
	$x = $WorkSheet.Columns.Item("A:A").EntireRow.AutoFit();

    $x = $WorkSheet.Cells.Item(2,"B").Select();
    $Excel.ActiveWindow.FreezePanes = $true;

	Write-Host "";
}

<#
.DESCRIPTION
	This function check the affinity name for profiles
#>
function HasAffinity($find, $prof) {

    $prof_values = $prof.ToLower().Split("-");

    for ($i = 0; $i -lt $prof_values.Length; $i++) {
        if ($find.ToLower().Replace("-", "").Contains($prof_values[$i])) {
            return $true;
        }
    }

    return $false;
}

<#
.DESCRIPTION
	This function create and fills worksheet for virtual server
#>
function Export-WorkSheetVS($WorkBook) {

	$WorkSheet = $WorkBook.WorkSheets.Add();
	$WorkSheet.Name = "Virtual Server";
    
    $row = Create-Header($WorkSheet);
	$WorkSheet.Cells.Item($row, $colVS_Name) = "Name";
	$WorkSheet.Cells.Item($row, $colVS_Source) = "Source";
	$WorkSheet.Cells.Item($row, $colVS_Destination) = "Destination";
	$WorkSheet.Cells.Item($row, $colVS_Port) = "Port";
	$WorkSheet.Cells.Item($row, $colVS_Protocol) = "Protocol";
	$WorkSheet.Cells.Item($row, $colVS_Description) = "Description";
	$WorkSheet.Cells.Item($row, $colVS_DefaultPool) = "Default pool";
	$WorkSheet.Cells.Item($row, $colVS_Persistence) = "Persistence";
	$WorkSheet.Cells.Item($row, $colVS_FallbackPersist) = "Fallback Persistence";
	$WorkSheet.Cells.Item($row, $colVS_SNAT) = "SNAT";
	$WorkSheet.Cells.Item($row, $colVS_SNATPool) = "SNAT Pool";
	$WorkSheet.Cells.Item($row, $colVS_VLAN) = "VLAN";
	$WorkSheet.Cells.Item($row, $colVS_ClientSSL) = "ClientSSL";
	$WorkSheet.Cells.Item($row, $colVS_ServerSSL) = "ServerSSL";
	$WorkSheet.Cells.Item($row, $colVS_TCPClient) = "TCP Client Profile";
	$WorkSheet.Cells.Item($row, $colVS_TCPServer) = "TCP Server Profile";
	$WorkSheet.Cells.Item($row, $colVS_HTTPProfile) = "HTTP/FTP Profile";
	$WorkSheet.Cells.Item($row, $colVS_OthProfile) = "Other Profile";
	$WorkSheet.Cells.Item($row, $colVS_iRule) = "iRule";
	$WorkSheet.Cells.Item($row, $colVS_Policy) = "Policies";
    $row++;
    $x = $WorkSheet.Cells.Item("$row","B").Select();
    $Excel.ActiveWindow.FreezePanes = $true;

	Write-Host "Retrieving profile type list...";
	$prof_list = Invoke-Get "/ltm/profile/?`$select=link";

    $profile_types = @();
    foreach ($prof in $prof_list.items) {
        $type = $prof.reference.link.Split("?")[0].Split("/");
        $profile_types += $type[$type.length -1];
    }

	Write-Host "Retrieving virtual server list...";
	$vs_list = Invoke-Get "/ltm/virtual/";

	Write-Host "Filling spreadsheet..." -NoNewline;
	$rng = $WorkBook.WorkSheets.Item("Pool").Range("A:A");

    $profiles_ready = @{};

	foreach ($vs in $vs_list.items) {

        Write-Verbose "Virtual: $vs";
	
		$vs_address_name = $vs.destination.Split(":")[0].Replace("/", "~");
		$vs_address_port = $vs.destination.Split(":")[1];
		$vs_address = Invoke-Get "/ltm/virtual-address/$vs_address_name/";

		$WorkSheet.Cells.Item($row, $colVS_Name) = $vs.fullPath;
		$WorkSheet.Cells.Item($row, $colVS_Source) = $vs.source;
		$WorkSheet.Cells.Item($row, $colVS_Destination) = "$($vs_address.address)/$(Get-MaskLength -mask $vs.mask)";
		$WorkSheet.Cells.Item($row, $colVS_Port) = $vs_address_port;
		$WorkSheet.Cells.Item($row, $colVS_Protocol) = $vs.ipProtocol;
		$WorkSheet.Cells.Item($row, $colVS_Description) = $vs.description;
		if ($vs.pool) {
			$find = $rng.Find($vs.pool);
			if ($find.Address() -ne "") {
				$lnk = $WorkSheet.Hyperlinks.Add($WorkSheet.Range($colVS_DefaultPool + $row), "", `
						"Pool!" + $find.Address(), $vs.pool, $vs.pool);
			}
		}
		
		$persistence = "";
		if ($vs.persist) {
			for ($i=0; $i -lt $vs.persist.count; $i++) {
				if ($persistence -ne "") { $persistence += "`r`n"; }
				$persistence += $vs.persist[$i].name;
			}
		}
		$WorkSheet.Cells.Item($row, $colVS_Persistence) = $persistence;
		$WorkSheet.Cells.Item($row, $colVS_FallbackPersist) = $vs.fallbackPersistence;
		
		$snat = "";
		if ($vs.sourceAddressTranslation) {
			$snat = $vs.sourceAddressTranslation[0].type;
		}
		
		$WorkSheet.Cells.Item($row, $colVS_SNAT) = $snat;
		if ($snat -eq "snat") {
			$WorkSheet.Cells.Item($row, $colVS_SNATPool) = $vs.sourceAddressTranslation[0].pool;
		}

        if ($vs.vlans) {
           $WorkSheet.Cells.Item($row, $colVS_VLAN) = [String]::Join("`r`n", $vs.vlans);
        }

		if ($vs.rules) {
            $WorkSheet.Cells.Item($row, $colVS_iRule) = [String]::Join("`r`n", $vs.rules);
		}
		
		if ($vs.profilesReference) {

            Write-Verbose "Checking profiles for vs: $vs.fullPath";
			$profiles = Invoke-Get "/ltm/virtual/$($vs.fullPath.Replace('/', '~'))/profiles/?`$select=fullPath,context";

			for ($i=0; $i -lt $profiles.items.count; $i++) {
				$find = $profiles.items[$i].fullPath.Replace("/", "~");

                Write-Verbose "Find profile: $find";
                $profile_list = @();
				#SMTP do not have in tmsh?
                if ($find -notmatch "~Common~smtp") {
                    $profile_list += $profile_types;
                }

                $retry = $false;
                $found = $false;
                for ($j=$profile_list.length - 1; $j -ge 0; $j--) {
                    
                    $prof = $profile_list[$j];

                    if (-not $retry) {
                        $hasAffinity = HasAffinity -find $find -prof $prof;
                    }
                    if ($retry -eq $true -or $hasAffinity -eq $true) {
                        try 
                        {
                            $type = $null;
                            if ($profiles_ready.ContainsKey($find)) {
                                $type = $profiles_ready[$find];
                            } else {
                                $s = Invoke-Get "/ltm/profile/$prof/$find/?`$select=kind,name,fullPath,defaultsFrom";
                                $type = $s.kind.Split(":")[3];
                                $profiles_ready.Add($find, $type);
                            }

                            switch ($type) {
                                {$_ -in "http", "ftp"} {
                                    if (-not [String]::IsNullOrEmpty($WorkSheet.Cells.Item($row, $colVS_HTTPProfile).Value2)) { $WorkSheet.Cells.Item($row, $colVS_HTTPProfile).Value2 += "`n" }
                                    $WorkSheet.Cells.Item($row, $colVS_HTTPProfile).Value2 += $profiles.items[$i].fullPath;
                                }
                                "tcp" {
                                    if ($profiles.items[$i].context -eq "serverside") {
                                        if (-not [String]::IsNullOrEmpty($WorkSheet.Cells.Item($row, $colVS_TCPServer).Value2)) { $WorkSheet.Cells.Item($row, $colVS_TCPServer).Value2 += "`n" }
                                        $WorkSheet.Cells.Item($row, $colVS_TCPServer).Value2 += $profiles.items[$i].fullPath;
                                    } else {
                                        if (-not [String]::IsNullOrEmpty($WorkSheet.Cells.Item($row, $colVS_TCPClient).Value2)) { $WorkSheet.Cells.Item($row, $colVS_TCPClient).Value2 += "`n" }
                                        $WorkSheet.Cells.Item($row, $colVS_TCPClient).Value2 += $profiles.items[$i].fullPath;
                                    }
                                }
                                "client-ssl" {
                                    if (-not [String]::IsNullOrEmpty($WorkSheet.Cells.Item($row, $colVS_HTTPProfile).Value2)) { $WorkSheet.Cells.Item($row, $colVS_HTTPProfile).Value2 += "`n" }
                                    $WorkSheet.Cells.Item($row, $colVS_ClientSSL).Value2 += $profiles.items[$i].fullPath;
                                }
                                "server-ssl" {
                                    if (-not [String]::IsNullOrEmpty($WorkSheet.Cells.Item($row, $colVS_ServerSSL).Value2)) { $WorkSheet.Cells.Item($row, $colVS_ServerSSL).Value2 += "`n" }
                                    $WorkSheet.Cells.Item($row, $colVS_ServerSSL).Value2 += $profiles.items[$i].fullPath;
                                }
                                default {
                                    if (-not [String]::IsNullOrEmpty($WorkSheet.Cells.Item($row, $colVS_OthProfile).Value2)) { $WorkSheet.Cells.Item($row, $colVS_OthProfile).Value2 += "`n" }
                                    $WorkSheet.Cells.Item($row, $colVS_OthProfile).Value2 += $profiles.items[$i].fullPath;
                                }
                            }

                            $found = $true;
                            break;
                        } catch {
                            #dummy
                            $profile_list = $profile_list -ne $prof;
                        }
                    }
                    if (($j -eq 0) -and (-not $retry)) {
                        $j = $profile_list.length - 1;
                        $retry = $true;
                    }
                }

                if (-not $found) {
                    if (-not [String]::IsNullOrEmpty($WorkSheet.Cells.Item($row, $colVS_OthProfile).Value2)) { $WorkSheet.Cells.Item($row, $colVS_OthProfile).Value2 += "`n" }
                    $WorkSheet.Cells.Item($row, $colVS_OthProfile).Value2 += $profiles.items[$i].fullPath;

                    if (-not $profiles_ready.ContainsKey($find)) {
                        $profiles_ready.Add($find, $null);
                    }

                    #SMTP or not provisioned
                   # Write-Host "";
                   # Write-Warning "Unknown type profile '$($profiles.items[$i].fullPath)'";
                }
			}
		}

        if ($vs.policiesReference.items) {
            for ($i=0; $i -lt $vs.policiesReference.items.count; $i++) {
                if (-not [String]::IsNullOrEmpty($WorkSheet.Cells.Item($row, $colVS_Policy).Value2)) { $WorkSheet.Cells.Item($row, $colVS_Policy).Value2 += "`n" }
                $WorkSheet.Cells.Item($row, $colVS_Policy).Value2 += $vs.policiesReference.items[$i].fullPath;
            }
        }

		Write-Host "." -NoNewline;
		$row++;
	}
    
    #A-T
    for ($i = 1; $i -le 20; $i++) {
	    $WorkSheet.Columns.Item($i).ColumnWidth = 50;
    }
	$x = $WorkSheet.Columns.Item("A:Z").EntireColumn.AutoFit();
	$x = $WorkSheet.Columns.Item("A:A").EntireRow.AutoFit();
    
    $x = $WorkSheet.Cells.Item(4,"B").Select();
    $Excel.ActiveWindow.FreezePanes = $true;

	Write-Host "";
}

<#
.DESCRIPTION
	This function export data to excel file
#>
function Export2Excel() {
	Try
	{
		Write-Host "`nStarting Excel Application...";
		$Excel = New-Object -ComObject Excel.Application;
		$WorkBook = $Excel.WorkBooks.Add();

		Export-WorkSheetNode -WorkBook $WorkBook;
		Export-WorkSheetPool -WorkBook $WorkBook;
		Export-WorkSheetVS -WorkBook $WorkBook;
		$date = (Get-Date -Format "MMddyyyy")
		#$parentFolder =   Split-Path -Parent $PSCommandPath
		$parentFolder ="\\sharecopter-ecm-ltd.cr.eurocopter.corp\sites\ecimoper\balancing\Shared Documents\Exploitation\Export F5"
		$filename = "BIGIP_`($($script:device.settings.hostname)`)-$date.xls";
		$WorkBook.SaveAs("$parentFolder\$filename");
	}
	Catch
	{
		$ErrorMessage = $_.Exception.Message;
		Write-Error $ErrorMessage;
	}
	Finally
	{
        Write-Host "Completed. File as been save here: $parentFolder\$filename";
		$Excel.Visible=$false;
        $Excel.Quit();
	}
}
#endregion

#region Initialization

if ( Initialize ) {
	Export2Excel;
}

#endregion