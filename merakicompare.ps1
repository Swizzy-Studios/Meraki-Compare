$key = read-host ("Please enter your API key or file name containing your encrypted key")
if (test-path $key){
	$k = get-content $key | convertto-securestring
	$ko = new-object System.management.automation.pscredential -argumentlist 'dummy',$k
	$apikey = $ko.getnetworkcredential().Password
}else{
	$apikey = $key
}
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.add('Authorization',('Bearer ' + $apikey))
$headers.add('Content-Type','application/json')
$ProgressPreference = 'SilentlyContinue'

function ConvertTo-IPv4MaskString ($inPrefix){
	$mask = [int]$inPrefix
	[IPaddress] $ip = 0;
	if ($mask -eq 0){
		$ip = [IPAddress] "0.0.0.0"
	}else{
		$ip.address = ([UInt32]::MaxValue) -shl (32 - $mask) -shr (32 - $mask)
	}
	return $ip
}

write-host "Loading Networks..."
try{
	$orgs = invoke-webrequest -headers $headers -method GET https://api.meraki.com/api/v1/organizations
}catch{
	"Error loading Organizations. Exiting"
}
$orgs = $orgs.content | convertfrom-json
$org = [ordered]@{
	ID = ""
	Company = ""
	Network = ""
	NetworkID = ""
	OrgID = ""
}
$allnets = @()
$c = 0
foreach ($o in $orgs){
	try{
		$nw = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/organizations/'+$o.id+'/networks')
	}catch{
		write-output ("Error loading network for " + $o.Name + ", " + $n.Name)
		continue
	}
	$nw = $nw.content | convertfrom-json
	foreach ($n in $nw){
		$temporg = [pscustomobject]$org
		$temporg.Company = $o.Name
		$temporg.Network = $n.Name
		$temporg.ID = $C
		$temporg.NetworkID = $n.ID
		$temporg.OrgID = $o.ID
		$c += 1
		$allnets += $temporg
	}
}
$allnets | select ID,Company,Network | ft
$netsel = read-host "Select which networks you would like to review access rules for. Separate by commas, use 'all' for all networks"
$netsel = $netsel -split ","
$groupsel = @()
if ($netsel -eq "all"){
	$groupsel = $allnets
}else{
	foreach ($s in $netsel){
		$groupsel += $allnets[[int]$s]
	}
}
$groupsel
$subnetobj = [ordered]@{
	subnet = ""
	name = ""
	applianceIP = ""
	networkType = ""
	arules = @()
	drules = @()
	vlan = ''
	orgid = ''
}
$ssidobj = [ordered]@{
	subnet = ""
	name = ""
	networkType = ""
	arules = @()
	drules = @()
	vlan = ''
	isguest = ''
}
$ruleexceptionsobj = [ordered]@{
	priority = ''
	name = ""
	action = ""
	direction = ""
	destination = ""
	dstip = ""
	dstsubnet = ''
	relevant = ''
}
$rulelistobj = [ordered]@{
	ip = ""
	subnet = ""
	priority = ""
}
$bwsubnetobj = [ordered]@{
	subnetIP = ""
	subnetmask = ""
}
$l3switchobj = [ordered]@{
	switchobj = @()
	subnets = @()
}
foreach ($n in $groupsel){
	$subnets = @()
	$e = 1
	$singleLan = try{invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/networks/'+$n.NetworkID+'/appliance/singleLan')}catch{$e = $_}
	if ($e -eq 1){
		$singleLan = $singleLan.content
		$singleLan = $singleLan | convertfrom-json
		$slansubnetobj = [pscustomobject]$subnetobj
		$slansubnetobj.subnet = $singleLan.subnet
		$slansubnetobj.name = "Default"
		$slansubnetobj.applianceIP = $singleLan.applianceIP
		$slansubnetobj.networkType = "singlelan"
		$slansubnetobj.orgid = $n.orgid
		$subnets += $slansubnetobj
	}elseif ($e | select-string "(400) Bad Request" -simple){
		$e2 = 1
		try {$vlans = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/networks/'+$n.NetworkID+'/appliance/vlans')}catch{$e2 = $_}
		if ($e2 -eq 1){
			$vlans = $vlans.content | convertfrom-json
			foreach ($v in $vlans){
				$vlansubnetobj = [pscustomobject]$subnetobj
				$vlansubnetobj.subnet = $v.subnet
				$vlansubnetobj.name = $v.name
				$vlansubnetobj.applianceIP = $v.applianceIP
				$vlansubnetobj.networkType = "vlan"
				$vlansubnetobj.vlan = $v.id
				$vlansubnetobj.orgid = $n.orgid
				$subnets += $vlansubnetobj
			}
		}elseif ($e2 | select-string "(400) Bad Request" -simple){
			write-output ("Error fetching subnets, exiting")
			exit(1)
		}
	}else{
		write-output ("Error fetching subnets, exiting")
		exit(1)
	}
	$devices = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/networks/'+$n.NetworkId+'/devices')
	$devices = $devices.content | convertfrom-json
	$routerserials = @()
	"Loading network devices..."
	foreach ($d in $devices){
		if (($d.wan1ip -ne $null) -or ($d.wan2ip -ne $null)){
			$routerserials += $d.serial
		}
	}
	'Calculating subnet masks for known networks...'
	$clientsubnets = @()
	foreach ($s in $subnets){
		$tempbwsubnetobj = [pscustomobject] $bwsubnetobj
		$sparts = $s.subnet.split("/")
		$tempbwsubnetobj.subnetIP = [ipaddress]$sparts[0]
		$tempbwsubnetobj.subnetmask = ConvertTo-IPv4MaskString($sparts[1])
		$clientsubnets += $tempbwsubnetobj
		$tempsubnetobj
	}
	'Loading client lists...'
	foreach ($s in $routerserials){
		
		$clients = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/devices/' + $s + '/clients')
		$clients = $clients.content | convertfrom-json
		foreach ($c in $clients){
			$tempip = [ipaddress] $c.ip
			$subnetfound = 0
			foreach ($cs in $clientsubnets){
				if ($cs.subnetIP.address -eq ($tempip.address -band $cs.subnetmask.address)){
					$subnetfound = 1
					break
				}
			}
			if ($subnetfound -eq 0){
				write-output ("Client " + $tempip.ipaddresstostring + " found not matching other known subnets. Defaulting to subnet of /24.")
				write-output ("Checking VLAN ID...")
				$tempsubnetobj = [pscustomobject]$subnetobj
				$tempbwsubnetobj = [pscustomobject]$bwsubnetobj
				if ($c.vlan -eq "0"){
					write-output ("VLAN ID 0 found. Possible VPN client, treating as such but this should be checked manually")
					$tempsubnetobj.networkType = "vpn"
				}else{
					write-output ("VLAN ID " + $c.vlan + " found. Adding to list")
					$tempsubnetobj.networkType = "unknown"
				}
				$ipparts = $c.ip.split(".")
				$tempsubnetobj.subnet = ($ipparts[0] + "." + $ipparts[1] + "." + $ipparts[2] + ".0/24")
				$tempsubnetobj.name = "Unknown"
				$tempsubnetobj.applianceIP = $ipparts[0] + "." + $ipparts[1] + "." + $ipparts[2] + ".0 *"
				$tempbwsubnetobj.subnetIP = [ipaddress] ($ipparts[0] + "." + $ipparts[1] + "." + $ipparts[2] + ".0")
				$tempbwsubnetobj.subnetmask = ConvertTo-IPv4MaskString("24")
				$clientsubnets += $tempbwsubnetobj
				$subnets += $tempsubnetobj
			}
		}
	}
	$l3switchfound = 0
	$switchrouters = @()
	foreach ($device in $devices){
		if ($device.firmware | select-string "switch"){
			$devicerouting = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/devices/' + $device.serial + '/switch/routing/interfaces')
			if ($devicerouting.content -ne '[]'){
				$l3switchfound = 1
				$devicerouting = $devicerouting.content | convertfrom-json
				$templ3switchobj = [pscustomobject]$l3switchobj
				$templ3switchobj.switchobj = $device
				$templ3switchobj.subnets = $devicerouting
				$switchrouters += $templ3switchobj
				foreach ($subnet in $devicerouting){
					$tempsubnetobj = [pscustomobject]$subnetobj
					$tempsubnetobj.name = $subnet.name
					$tempsubnetobj.subnet = $subnet.subnet
					$tempsubnetobj.applianceIP = $subnet.interfaceIp
					$tempsubnetobj.networkType = 'L3 Switch'
					$tempsubnetobj.vlan = $subnet.vlanID
					$tempsubnetobj.orgid = $n.orgid
					$subnets += $tempsubnetobj
				}
			}
		}
	}
	$outboundrules = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/networks/' + $n.Networkid + '/appliance/firewall/l3FirewallRules')
	$inboundrules = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/networks/' + $n.networkid + '/appliance/firewall/inboundFirewallRules')
	$outboundrules = $outboundrules.content
	$outboundrules = $outboundrules | convertfrom-json
	$outboundrules = $outboundrules.rules
	$inboundrules = $inboundrules.content
	$inboundrules = $inboundrules | convertfrom-json
	$inboundrules = $inboundrules.rules
	
	foreach ($s in $subnets){
		$sparts = $s.subnet.split('/')
		$tempip = [ipaddress]$sparts[0]
		$tempsubnet = ConvertTo-IPv4MaskString($sparts[1])
		$denylist = @()
		$allowlist = @()
		$c = 0
		$grouped = 0
		foreach ($obr in $outboundrules){
			$srcnetparts = $obr.srccidr.split(',')
			foreach ($src in $srcnetparts){
				if ($src -eq "Any"){
					$tempsrcnet = [ipaddress] "0.0.0.0"
					$tempsrcsub = [ipaddress] "0.0.0.0"
				}elseif($src | select-string "VLAN(" -simple){
					$vlanregex = [regex]"\((.*?)\)"
					$vlanid = $src -match $vlanregex
					if ($vlanid -eq $true){
						$vlanid = $matches[1]
					}else{
						write-output ("Error identifying VLANID")
						continue
					}
					foreach ($s2 in $subnets){
						if ($s2.vlan -eq $vlanid){
							$srcparts = $s2.subnet.split("/")
							$tempsrcnet = [ipaddress] $srcparts[0]
							$tempsrcsub = ConvertTo-IPv4MaskString($srcparts[1])
						}
					}
				}elseif($src | select-string "GRP(" -simple){
					$grouped = 1
					$grpregex = [regex]"\((.*?)\)"
					$grpid = $src -match $grpregex
					if ($grpid -eq $true){
						$grpid = $matches[1]
					}else{
						write-output ("Error identifying GRPID")
						continue
					}
					try{
						$groupitems = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/organizations/' + $s.orgid + '/policyObjects/groups/' + $grpid)
					}catch{
						write-output ("Error Fetching group ID: " + $grpid)
						continue
					}
					$groupitems = $groupitems.content
					$groupitems = $groupitems | convertfrom-json
					foreach ($oid in $groupitems.objectids){

						$objnet = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/organizations/' + $s.orgid + '/policyObjects/' + $oid)
						$objnet = $objnet.content
						$objnet = $objnet | convertfrom-json
						$srcparts = $objnet.cidr.split("/")
						$tempsrcnet = [ipaddress] $srcparts[0]
						$tempsrcsub = ConvertTo-IPv4MaskString($srcparts[1])
						if ($tempsrcnet.address -eq ($tempip.address -band $tempsrcsub.address)){
					
							$dstnetparts = $obr.destcidr.split(",")
							$tempdestinations = @()
							foreach ($dst in $dstnetparts){
								$groupeddst = 0
								$tempdestination = [pscustomobject]$bwsubnetobj
								if ($dst -eq "Any"){
									$tempdestination.subnetip = [ipaddress] "0.0.0.0"
									$tempdestination.subnetmask = [ipaddress] "0.0.0.0"
									

								}elseif($dst | select-string "VLAN(" -simple){
									$vlanregex = [regex]"\((.*?)\)"
									$vlanid = $dst -match $vlanregex
									if ($vlanid -eq $true){
										$vlanid = $matches[1]
									}else{
										write-output ("Error identifying VLANID")
										continue
									}
									foreach ($s2 in $subnets){
										if ($s2.vlan -eq $vlanid){
											$dstparts = $s2.subnet.split("/")
											$tempdestination.subnetip = [ipaddress] $dstparts[0]
											$tempdestination.subnetmask = ConvertTo-IPv4MaskString($dstparts[1])
										}
									}
								}elseif($dst | select-string "GRP(" -simple){
									$groupeddst = 1
									$grpregex = [regex]"\((.*?)\)"
									$grpdstid = $dst -match $grpregex
									if ($grpdstid -eq $true){
										$grpdstid = $matches[1]
									}else{
										write-output ("Error identifying destination GRPID")
										continue
									}
									try{
										$groupdstitems = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/organizations/' + $s.orgid + '/policyObjects/groups/' + $grpdstid)
									}catch{
										write-output ("Error Fetching group ID: " + $grpdstid)
										continue
									}
									$groupdstitems = $groupdstitems.content
									$groupdstitems = $groupdstitems | convertfrom-json
									foreach ($oiddst in $groupdstitems.objectids){
										$objnetdst = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/organizations/' + $s.orgid + '/policyObjects/' + $oiddst)
										$objnetdst = $objnetdst.content
										$objnetdst = $objnetdst | convertfrom-json
										$dstarts = $objnetdst.cidr.split("/")
										$tempdestination.subnetip = [ipaddress] $dstparts[0]
										$tempdestination.subnetmask = convertto-ipv4maskstring($dstparts[1])
										$tempdestinations += $tempdestination
										
										if ($obr.policy -eq "allow"){
											$tempruleobj = [pscustomobject] $ruleexceptionsobj
											$tempruleobj.priority = $c
											$tempruleobj.dstip = $tempdestination.subnetip
											$tempruleobj.dstsubnet = $tempdestination.subnetmask
											$tempruleobj.direction = 'out'
											$tempruleobj.name = $obr.comment
											$tempruleobj.action = $obr.policy
											$tempruleobj.destination = $dst
											$allowlist += $tempruleobj
										}elseif($obr.policy -eq "deny"){
											$tempruleobj = [pscustomobject] $ruleexceptionsobj
											$tempruleobj.priority = $c
											$tempruleobj.dstip = $tempdestination.subnetip
											$tempruleobj.dstsubnet = $tempdestination.subnetmask
											$tempruleobj.direction = 'out'
											$tempruleobj.name = $obr.comment
											$tempruleobj.action = $obr.policy
											$tempruleobj.destination = $dst
											$denylist += $tempruleobj
										}
									}
								}elseif ($dst | select-string "OBJ(" -simple){
									$objregex = [regex]"\((.*?)\)"
									$objdstid = $dst -match $grpregex
									if ($objdstid -eq $true){
										$objdstid = $matches[1]
									}else{
										write-output ("Error identifying destination OBJID")
										continue
									}
									$objnetdst = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/organizations/' + $s.orgid + '/policyObjects/' + $objdstid)
									$objnetdst = $objnetdst.content
									$objnetdst = $objnetdst | convertfrom-json
									$dstparts = $objnetdst.cidr.split("/")
									$tempdestination.subnetip = [ipaddress] $dstparts[0]
									$tempdestination.subnetmask = convertto-ipv4maskstring($dstparts[1])
									$tempdestinations += $tempdestination
									
									if ($obr.policy -eq "allow"){
										$tempruleobj = [pscustomobject] $ruleexceptionsobj
										$tempruleobj.priority = $c
										$tempruleobj.dstip = $tempdestination.subnetip
										$tempruleobj.dstsubnet = $tempdestination.subnetmask
										$tempruleobj.direction = 'out'
										$tempruleobj.name = $obr.comment
										$tempruleobj.action = $obr.policy
										$tempruleobj.destination = $dst
										$allowlist += $tempruleobj
									}elseif($obr.policy -eq "deny"){
										$tempruleobj = [pscustomobject] $ruleexceptionsobj
										$tempruleobj.priority = $c
										$tempruleobj.dstip = $tempdestination.subnetip
										$tempruleobj.dstsubnet = $tempdestination.subnetmask
										$tempruleobj.direction = 'out'
										$tempruleobj.name = $obr.comment
										$tempruleobj.action = $obr.policy
										$tempruleobj.destination = $dst
										$denylist += $tempruleobj
									}
								}else{
									$dstparts = $dst.split("/")
									$tempdestination.subnetip = [ipaddress] $dstparts[0]
									$tempdestination.subnetmask = ConvertTo-IPv4MaskString($dstparts[1])
								}
								if ($groupeddst -eq 0){
									$tempdestinations += $tempdestination
								
								
									if ($obr.policy -eq "allow"){
										$tempruleobj = [pscustomobject] $ruleexceptionsobj
										$tempruleobj.priority = $c
										$tempruleobj.dstip = $tempdestination.subnetip
										$tempruleobj.dstsubnet = $tempdestination.subnetmask
										$tempruleobj.direction = 'out'
										$tempruleobj.name = $obr.comment
										$tempruleobj.action = $obr.policy
										$tempruleobj.destination = $dst
										$allowlist += $tempruleobj
									}elseif($obr.policy -eq "deny"){
										$tempruleobj = [pscustomobject] $ruleexceptionsobj
										$tempruleobj.priority = $c
										$tempruleobj.dstip = $tempdestination.subnetip
										$tempruleobj.dstsubnet = $tempdestination.subnetmask
										$tempruleobj.direction = 'out'
										$tempruleobj.name = $obr.comment
										$tempruleobj.action = $obr.policy
										$tempruleobj.destination = $dst
										$denylist += $tempruleobj
									}
								}
							}
						}

					}
				}elseif ($dst | select-string "OBJ(" -simple){
					$objregex = [regex]"\((.*?)\)"
					$objdstid = $dst -match $grpregex
					if ($objdstid -eq $true){
						$objdstid = $matches[1]
					}else{
						write-output ("Error identifying destination OBJID")
						continue
					}
					$objnetdst = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/organizations/' + $s.orgid + '/policyObjects/' + $objdstid)
					$objnetdst = $objnetdst.content
					$objnetdst = $objnetdst | convertfrom-json
					$dstparts = $objnetdst.cidr.split("/")
					$tempdestination.subnetip = [ipaddress] $dstparts[0]
					$tempdestination.subnetmask = convertto-ipv4maskstring($dstparts[1])
					$tempdestinations += $tempdestination
					
					if ($obr.policy -eq "allow"){
						$tempruleobj = [pscustomobject] $ruleexceptionsobj
						$tempruleobj.priority = $c
						$tempruleobj.dstip = $tempdestination.subnetip
						$tempruleobj.dstsubnet = $tempdestination.subnetmask
						$tempruleobj.direction = 'out'
						$tempruleobj.name = $obr.comment
						$tempruleobj.action = $obr.policy
						$tempruleobj.destination = $dst
						$allowlist += $tempruleobj
					}elseif($obr.policy -eq "deny"){
						$tempruleobj = [pscustomobject] $ruleexceptionsobj
						$tempruleobj.priority = $c
						$tempruleobj.dstip = $tempdestination.subnetip
						$tempruleobj.dstsubnet = $tempdestination.subnetmask
						$tempruleobj.direction = 'out'
						$tempruleobj.name = $obr.comment
						$tempruleobj.action = $obr.policy
						$tempruleobj.destination = $dst
						$denylist += $tempruleobj
					}
				}else{

					$srcparts = $src.split("/")
					$tempsrcnet = [ipaddress] $srcparts[0]
					$tempsrcsub = ConvertTo-IPv4MaskString($srcparts[1])
				}

				if ($grouped -eq 0){
					
					if ($tempsrcnet.address -eq ($tempip.address -band $tempsrcsub.address)){
						
						$dstnetparts = $obr.destcidr.split(",")
						$tempdestinations = @()
						foreach ($dst in $dstnetparts){
							$groupeddst = 0
							$tempdestination = [pscustomobject]$bwsubnetobj
							if ($dst -eq "Any"){
								$tempdestination.subnetip = [ipaddress] "0.0.0.0"
								$tempdestination.subnetmask = [ipaddress] "0.0.0.0"

							}elseif($dst | select-string "VLAN(" -simple){
								$vlanregex = [regex]"\((.*?)\)"
								$vlanid = $dst -match $vlanregex
								if ($vlanid -eq $true){
									$vlanid = $matches[1]
								}else{
									write-output ("Error identifying VLANID")
									continue
								}
								foreach ($s2 in $subnets){
									if ($s2.vlan -eq $vlanid){
										$dstparts = $s2.subnet.split("/")
										$tempdestination.subnetip = [ipaddress] $dstparts[0]
										$tempdestination.subnetmask = ConvertTo-IPv4MaskString($dstparts[1])
									}
								}
							}elseif($dst | select-string "GRP(" -simple){
								$groupeddst = 1
								$grpregex = [regex]"\((.*?)\)"
								$grpdstid = $dst -match $grpregex
								if ($grpdstid -eq $true){
									$grpdstid = $matches[1]
								}else{
									write-output ("Error identifying destination GRPID")
									continue
								}
								try{
									$groupdstitems = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/organizations/' + $s.orgid + '/policyObjects/groups/' + $grpdstid)
								}catch{
									write-output ("Error Fetching group ID: " + $grpdstid)
									continue
								}
								$groupdstitems = $groupdstitems.content
								$groupdstitems = $groupdstitems | convertfrom-json
								foreach ($oiddst in $groupdstitems.objectids){
									$objnetdst = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/organizations/' + $s.orgid + '/policyObjects/' + $oiddst)
									$objnetdst = $objnetdst.content
									$objnetdst = $objnetdst | convertfrom-json
									$dstparts = $objnetdst.cidr.split("/")
									$tempdestination.subnetip = [ipaddress] $dstparts[0]
									$tempdestination.subnetmask = convertto-ipv4maskstring($dstparts[1])
									$tempdestinations += $tempdestination
									
									if ($obr.policy -eq "allow"){
										$tempruleobj = [pscustomobject] $ruleexceptionsobj
										$tempruleobj.priority = $c
										$tempruleobj.dstip = $tempdestination.subnetip
										$tempruleobj.dstsubnet = $tempdestination.subnetmask
										$tempruleobj.direction = 'out'
										$tempruleobj.name = $obr.comment
										$tempruleobj.action = $obr.policy
										$tempruleobj.destination = $dst
										$allowlist += $tempruleobj
									}elseif($obr.policy -eq "deny"){
										$tempruleobj = [pscustomobject] $ruleexceptionsobj
										$tempruleobj.priority = $c
										$tempruleobj.dstip = $tempdestination.subnetip
										$tempruleobj.dstsubnet = $tempdestination.subnetmask
										$tempruleobj.direction = 'out'
										$tempruleobj.name = $obr.comment
										$tempruleobj.action = $obr.policy
										$tempruleobj.destination = $dst
										$denylist += $tempruleobj
									}
								}
							}elseif ($dst | select-string "OBJ(" -simple){
								$objregex = [regex]"\((.*?)\)"
								$objdstid = $dst -match $grpregex
								if ($objdstid -eq $true){
									$objdstid = $matches[1]
								}else{
									write-output ("Error identifying destination OBJID")
									continue
								}
								$objnetdst = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/organizations/' + $s.orgid + '/policyObjects/' + $objdstid)
								$objnetdst = $objnetdst.content
								$objnetdst = $objnetdst | convertfrom-json
								$dstparts = $objnetdst.cidr.split("/")
								$tempdestination.subnetip = [ipaddress] $dstparts[0]
								$tempdestination.subnetmask = convertto-ipv4maskstring($dstparts[1])
								$tempdestinations += $tempdestination
								
								if ($obr.policy -eq "allow"){
									$tempruleobj = [pscustomobject] $ruleexceptionsobj
									$tempruleobj.priority = $c
									$tempruleobj.dstip = $tempdestination.subnetip
									$tempruleobj.dstsubnet = $tempdestination.subnetmask
									$tempruleobj.direction = 'out'
									$tempruleobj.name = $obr.comment
									$tempruleobj.action = $obr.policy
									$tempruleobj.destination = $dst
									$allowlist += $tempruleobj
								}elseif($obr.policy -eq "deny"){
									$tempruleobj = [pscustomobject] $ruleexceptionsobj
									$tempruleobj.priority = $c
									$tempruleobj.dstip = $tempdestination.subnetip
									$tempruleobj.dstsubnet = $tempdestination.subnetmask
									$tempruleobj.direction = 'out'
									$tempruleobj.name = $obr.comment
									$tempruleobj.action = $obr.policy
									$tempruleobj.destination = $dst
									$denylist += $tempruleobj
								}
							}else{
								$dstparts = $dst.split("/")
								$tempdestination.subnetip = [ipaddress] $dstparts[0]
								$tempdestination.subnetmask = ConvertTo-IPv4MaskString($dstparts[1])
							}
							
							if ($groupeddst -eq 0){
								$tempdestinations += $tempdestination
								if ($obr.policy -eq "allow"){
									$tempruleobj = [pscustomobject] $ruleexceptionsobj
									$tempruleobj.priority = $c
									$tempruleobj.dstip = $tempdestination.subnetip
									$tempruleobj.dstsubnet = $tempdestination.subnetmask
									$tempruleobj.direction = 'out'
									$tempruleobj.name = $obr.comment
									$tempruleobj.action = $obr.policy
									$tempruleobj.destination = $dst
									$allowlist += $tempruleobj
								}elseif($obr.policy -eq "deny"){
									$tempruleobj = [pscustomobject] $ruleexceptionsobj
									$tempruleobj.priority = $c
									$tempruleobj.dstip = $tempdestination.subnetip
									$tempruleobj.dstsubnet = $tempdestination.subnetmask
									$tempruleobj.direction = 'out'
									$tempruleobj.name = $obr.comment
									$tempruleobj.action = $obr.policy
									$tempruleobj.destination = $dst
									$denylist += $tempruleobj
								}
							}
						}
					}
				}
			}
			$c+=1
		}


		foreach ($allowrule in $allowlist){
			$s.arules += $allowrule
		}
		foreach ($denyrule in $denylist){
			$s.drules += $denyrule
		}
		
		
	}

	Write-output("================================================================================")
	Write-output("FIREWALL ACCESS RULES")
	Write-output('--------------------------------------------------------------------------------')
	foreach ($s in $subnets){
		$s | select name,subnet,networktype,vlan | ft

		foreach ($s2 in $subnets){
			$rulecompareallow = ''
			$rulecomparedeny = ''
			$sparts = $s2.subnet.split('/')
			$tempip = [ipaddress]$sparts[0]
			$tempsubnet = ConvertTo-IPv4MaskString($sparts[1])

			foreach ($ar in $s.arules){

				if ($ar.dstip.address -eq ($tempip.address -band $ar.dstsubnet.address)){
					if ($ar.priority -lt $rulecompareallow.priority){
						$rulecompareallow = $ar
					}elseif($rulecompareallow -eq ''){
						$rulecompareallow = $ar
					}
				}
			}
			
			foreach ($ad in $s.drules){

				if ($ad.dstip.address -eq ($tempip.address -band $ad.dstsubnet.address)){
					if ($ad.priority -lt $rulecompareallow.priority){
						$rulecomparedeny = $ad
					}elseif($rulecomparedeny = ''){
						$rulecomparedeny = $ad
					}
				}
			}
			if ($rulecompareallow.priority -lt $rulecomparedeny.priority){
				write-output ("      |_ Has access to " + $s2.subnet + " [" + $s2.Name + "] through rule " + $rulecompareallow.name)
			}elseif (($rulecompareallow -ne '') -and ($rulecomparedeny -eq '')){
				write-output ("      |_ Has access to " + $s2.subnet + " [" + $s2.Name + "] through rule " + $rulecompareallow.name)
			}

		
		}
	}
	Write-output("================================================================================")
	write-output("")
	Write-output("================================================================================")
	write-output("L3 SWITCH ACCESS CONTROL")
	write-output("If the source subnet is not routed by the L3 switch, then the firewall rules will apply")
	write-output("The following comparison only considers the ACL rules on the L3 switch")
	Write-output('--------------------------------------------------------------------------------')

	if ($l3switchfound -ne 0){
		$switchacl = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/networks/' + $n.networkid + '/switch/accessControlLists')
		$switchacl = $switchacl.content | convertfrom-json
		$switchacl = $switchacl.rules
		$subnetscopy = $subnets
		foreach ($s in $subnetscopy){
			$allowlist = @()
			$denylist = @()
			$s.arules = @()
			$s.drules = @()
			$sparts = $s.subnet.split('/')
			$tempip = [ipaddress]$sparts[0]
			$tempsubnet = ConvertTo-IPv4MaskString($sparts[1])
			$csw = 0
			foreach ($aclr in $switchacl){
				if ($aclr.srccidr -eq 'any'){
					$tempsrcnet = [ipaddress] "0.0.0.0"
					$tempsrcsub = [ipaddress] "0.0.0.0"
				}else{
					$srcparts = $aclr.srccidr.split("/")
					$tempsrcnet = [ipaddress] $srcparts[0]
					$tempsrcsub = ConvertTo-IPv4MaskString($srcparts[1])
				}

				if ($tempsrcnet.address -eq ($tempip.address -band $tempsrcsub.address)){
					$dst = $aclr.dstCidr
					if ($dst -eq "any"){
						$tempdstnet = [ipaddress] "0.0.0.0"
						$tempdstsub = [ipaddress] "0.0.0.0"
					}else{
						$dstparts = $dst.split('/')
						$tempdstnet = [ipaddress] $dstparts[0]
						$tempdstsub = ConvertTo-IPv4MaskString($dstparts[1])
					}
					if ($aclr.policy -eq "allow"){
						$tempruleobj = [pscustomobject] $ruleexceptionsobj
						$tempruleobj.priority = $csw
						$tempruleobj.dstip = $tempdstnet
						$tempruleobj.dstsubnet = $tempdstsub
						$tempruleobj.direction = 'out'
						$tempruleobj.name = $aclr.comment
						$tempruleobj.action = $aclr.policy
						$tempruleobj.destination = $dst
						$allowlist += $tempruleobj
					}elseif($aclr.policy -eq "deny"){
						$tempruleobj = [pscustomobject] $ruleexceptionsobj
						$tempruleobj.priority = $csw
						$tempruleobj.dstip = $tempdstnet
						$tempruleobj.dstsubnet = $tempdstsub
						$tempruleobj.direction = 'out'
						$tempruleobj.name = $aclr.comment
						$tempruleobj.action = $aclr.policy
						$tempruleobj.destination = $dst
						$denylist += $tempruleobj
					}
				}
				$csw += 1
			}
			foreach ($allowrule in $allowlist){
				$s.arules += $allowrule
			}
			foreach ($denyrule in $denylist){
				$s.drules += $denyrule
			}
			
		}
		$swc = 1
		foreach ($sw in $switchrouters){
			Write-output('--------------------------------------------------------------------------------')
			write-output('L3 Switch: ' + $sw.switchobj.Name + ' Switch Model: ' + $sw.switchobj.Model + ' Switch IP: ' + $sw.switchobj.LanIp)
			Write-output('--------------------------------------------------------------------------------')
			foreach ($s in $subnetscopy){
				$s | select name,subnet,networktype,vlan | ft
				foreach ($s2 in $sw.subnets){
					$rulecompareallow = ''
					$rulecomparedeny = ''
					$sparts = $s2.subnet.split('/')
					$tempip = [ipaddress]$sparts[0]
					$tempsubnet = ConvertTo-IPv4MaskString($sparts[1])

					foreach ($ar in $s.arules){
						
						if ($ar.dstip.address -eq ($tempip.address -band $ar.dstsubnet.address)){
							if ($ar.priority -lt $rulecompareallow.priority){
								$rulecompareallow = $ar
							}elseif($rulecompareallow -eq ''){
								$rulecompareallow = $ar
							}
						}
					}
					
					foreach ($ad in $s.drules){

						if ($ad.dstip.address -eq ($tempip.address -band $ad.dstsubnet.address)){
							if ($ad.priority -lt $rulecompareallow.priority){
								$rulecomparedeny = $ad
							}elseif($rulecomparedeny = ''){
								$rulecomparedeny = $ad
							}
						}
					}
					if ($rulecompareallow.priority -lt $rulecomparedeny.priority){
						write-output ("      |_ Has access to " + $s2.subnet + " [" + $s2.Name + "] through rule " + $rulecompareallow.name)
						foreach ($s3 in $subnetscopy){
							if (($s3.subnet -eq $s2.subnet) -and ($s3.networktype -eq 'L3 Switch')){
								#top
								$rulecompareallow2 = ''
								$rulecomparedeny2 = ''
								$sparts2 = $s.subnet.split('/')
								$tempip2 = [ipaddress]$sparts2[0]
								$tempsubnet2 = ConvertTo-IPv4MaskString($sparts2[1])

								foreach ($ar in $s3.arules){
									
									if ($ar.dstip.address -eq ($tempip2.address -band $ar.dstsubnet.address)){
										if ($ar.priority -lt $rulecompareallow2.priority){
											$rulecompareallow2 = $ar
										}elseif($rulecompareallow2 -eq ''){
											$rulecompareallow2 = $ar
										}
									}
								}
								
								foreach ($ad in $s3.drules){

									if ($ad.dstip.address -eq ($tempip2.address -band $ad.dstsubnet.address)){
										if ($ad.priority -lt $rulecompareallow2.priority){
											$rulecomparedeny2 = $ad
										}elseif($rulecomparedeny2 = ''){
											$rulecomparedeny2 = $ad
										}
									}
								}
								if ($rulecompareallow2.priority -lt $rulecomparedeny2.priority){
									write-output ("          |_ Stateless inverse packets from " + $s3.subnet + " [" + $s3.Name + "] ARE allowed through rule " + $rulecompareallow2.name)
								}elseif(($rulecompareallow2 -ne '') -and ($rulecomparedeny2 -eq '')){
									write-output ("          |_ Stateless inverse packets from " + $s3.subnet + " [" + $s3.Name + "] ARE allowed through rule " + $rulecompareallow2.name)
								}else{
									write-output ("          |_ Stateless inverse packets from " + $s3.subnet + " [" + $s3.Name + "] are NOT allowed")

								}
							}
						}
								#bottom
					}elseif (($rulecompareallow -ne '') -and ($rulecomparedeny -eq '')){
						write-output ("      |_ Has access to " + $s2.subnet + " [" + $s2.Name + "] through rule " + $rulecompareallow.name)
						foreach ($s3 in $subnetscopy){
							if (($s3.subnet -eq $s2.subnet) -and ($s3.networktype -eq 'L3 Switch')){
								
								#top
								$rulecompareallow2 = ''
								$rulecomparedeny2 = ''
								$sparts2 = $s.subnet.split('/')
								$tempip2 = [ipaddress]$sparts2[0]
								$tempsubnet2 = ConvertTo-IPv4MaskString($sparts2[1])

								foreach ($ar in $s3.arules){
									
									if ($ar.dstip.address -eq ($tempip2.address -band $ar.dstsubnet.address)){
										if ($ar.priority -lt $rulecompareallow2.priority){
											$rulecompareallow2 = $ar
										}elseif($rulecompareallow2 -eq ''){
											$rulecompareallow2 = $ar
										}
									}
								}
								
								foreach ($ad in $s3.drules){

									if ($ad.dstip.address -eq ($tempip2.address -band $ad.dstsubnet.address)){
										if ($ad.priority -lt $rulecompareallow2.priority){
											$rulecomparedeny2 = $ad
										}elseif($rulecomparedeny2 = ''){
											$rulecomparedeny2 = $ad
										}
									}
								}
								if ($rulecompareallow2.priority -lt $rulecomparedeny2.priority){
									write-output ("          |_ Stateless inverse packets from " + $s3.subnet + " [" + $s3.Name + "] ARE allowed through rule " + $rulecompareallow2.name)
								}elseif(($rulecompareallow2 -ne '') -and ($rulecomparedeny2 -eq '')){
									write-output ("          |_ Stateless inverse packets from " + $s3.subnet + " [" + $s3.Name + "] ARE allowed through rule " + $rulecompareallow2.name)
								}else{
									write-output ("          |_ Stateless inverse packets from " + $s3.subnet + " [" + $s3.Name + "] are NOT allowed")
								}
							}
						}
					}
				}
			}
			
		}
	}else{
		write-output("No L3 Switches found")
	}
				


	Write-output("================================================================================")
	write-output("")
	Write-output("================================================================================")
	write-output("WIRELESS ACCESS RULES")
	write-output("Traffic is restricted by the APs with the following ruleset.")
	write-output("If traffic is permitted and is routed through the firewall or L3 switch, those rules will apply")
	Write-output('--------------------------------------------------------------------------------')
	#Wireless Firewall Rules
	try{
		$ssids = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/networks/' + $n.networkid + '/wireless/ssids')
	}catch{
		write-output("Error gathering SSID information, do they have wireless?")
		continue
	}
	$ssids = $ssids.content
	$ssids = $ssids | convertfrom-json
	$enabledssid = @()
	$skipssid = 0
	foreach ($ssid in $ssids){
		if ($ssid.enabled -eq "True"){
			$enabledssid += $ssid
		}
	}
	$ssidlist = @()
	foreach ($ssid in $enabledssid){
		$tempssidobj = [pscustomobject]$ssidobj
		$tempssidobj.name = $ssid.Name
		$tempssidobj.networktype = 'WiFi'
		if ($ssid.ipAssignmentMode -eq 'Bridge mode'){
			$tempssidobj.isguest = $false
		}elseif ($ssid.ipAssignmentMode -eq 'NAT mode'){
			$tempssidobj.isguest = $true
		}
		if ($ssid.useVlanTagging -eq 'True'){
			$vlanfound = 0
			$tempssidobj.vlan = $ssid.defaultVlanId
			foreach ($s in $subnets){
				if ($s.vlan -eq $ssid.defaultVlanId){
					$vlanfound = 1
					$tempssidobj.subnet = $s.subnet
				}
			}
			if ($vlanfound -eq 0){
				write-output("VLAN Not found " + $ssid.vlan + ' for ' + $ssid.name)
				if ($tempssidobj.isguest = $true){
					write-output ("Detected guest network, defaulting to 10.0.0.0/8")
					$tempssidobj.subnet = '10.0.0.0/8'
				}else{
					write-output ("No VLAN found in network, is there another DHCP server? Skipping SSID...")
					$skipssid = 1
					continue
				}
			}
		}elseif ($ssid.ipAssignmentMode -eq 'Bridge mode'){
			$defaultfound = 0
			foreach($s in $subnets){
				if ($s.Name -eq 'Default'){
					$defaultfound = 1
					$tempssidobj.subnet = $s.subnet
					$tempssidobj.vlan = '1'
				}
			}
			if ($defaultfound -eq 0){
				$skipssid = 1
			}
		}
		if ($skipssid -eq 0){
			$srcnetparts = $tempssidobj.subnet.split('/')
			$tempsrcnet = $srcnetparts[0]
			$tempsrcsub = ConvertTo-IPv4MaskString($srcnetparts[1])
			$ssidrules = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/networks/' + $n.NetworkID + '/wireless/ssids/' + $ssid.number + '/firewall/l3FirewallRules')
			$ssidrules = $ssidrules.content
			$ssidrules = $ssidrules | convertfrom-json
			$ssidrules = ($ssidrules).Rules
			$allowlist = @()
			$denylist = @()
			$cs = 0
			foreach ($srule in $ssidrules){
				$grouped = 0
				$dstnetparts = $srule.Destcidr.split(',')
				$tempdestinations = @()
				foreach ($dst in $dstnetparts){
					$tempdestination = [pscustomobject]$bwsubnetobj
					if ($dst -eq "Any"){
						$tempdestination.subnetip = [ipaddress] "0.0.0.0"
						$tempdestination.subnetmask = [ipaddress] "0.0.0.0"
					}elseif($dst -eq "Local LAN"){
						$grouped = 1
						$tempdestLL = @()
						$tempdestination.subnetip = [ipaddress] "192.168.0.0"
						$tempdestination.subnetmask = [ipaddress] "255.255.0.0"
						$tempdestLL += $tempdestination
						$tempdestination = [pscustomobject]$bwsubnetobj
						$tempdestination.subnetip = [ipaddress] "10.0.0.0"
						$tempdestination.subnetmask = [ipaddress] "255.0.0.0"
						$tempdestLL += $tempdestination
						$tempdestination = [pscustomobject]$bwsubnetobj
						$tempdestination.subnetip = [ipaddress] "172.16.0.0"
						$tempdestination.subnetmask = [ipaddress] "255.240.0.0"
						$tempdestLL += $tempdestination
						
						foreach ($tempdst in $tempdestLL){
							$tempdestinations += $tempdst
							if ($srule.policy -eq "allow"){
								$tempruleobj = [pscustomobject] $ruleexceptionsobj
								$tempruleobj.priority = $cs
								$tempruleobj.dstip = $tempdst.subnetip
								$tempruleobj.dstsubnet = $tempdst.subnetmask
								$tempruleobj.direction = 'out'
								$tempruleobj.name = $srule.comment
								$tempruleobj.action = $srule.policy
								if ($tempdestLL.dstip -eq [ipaddress]"192.168.0.0"){
									$tempruleobj.destination = '192.168.0.0/16'
								}elseif($tempdestLL.dstip -eq [ipaddress]"10.0.0.0"){
									$tempruleobj.destination = '10.0.0.0/8'
								}elseif($tempdestLL.dstip -eq [ipaddress]"172.16.0.0"){
									$tempruleobj.destination = '172.16.0.0/12'
								}
								$allowlist += $tempruleobj
							}elseif($srule.policy -eq "deny"){
								$tempruleobj = [pscustomobject] $ruleexceptionsobj
								$tempruleobj.priority = $cs
								$tempruleobj.dstip = $tempdst.subnetip
								$tempruleobj.dstsubnet = $tempdst.subnetmask
								$tempruleobj.direction = 'out'
								$tempruleobj.name = $srule.comment
								$tempruleobj.action = $srule.policy
								if ($tempdestLL.dstip -eq [ipaddress]"192.168.0.0"){
									$tempruleobj.destination = '192.168.0.0/16'
								}elseif($tempdestLL.dstip -eq [ipaddress]"10.0.0.0"){
									$tempruleobj.destination = '10.0.0.0/8'
								}elseif($tempdestLL.dstip -eq [ipaddress]"172.16.0.0"){
									$tempruleobj.destination = '172.16.0.0/12'
								}
								$denylist += $tempruleobj
							}
						}
					}else{
						$dstparts = $dst.split("/")
						$tempdestination.subnetip = [ipaddress] $dstparts[0]
						$tempdestination.subnetmask = ConvertTo-IPv4MaskString($dstparts[1])
					}
					if ($grouped -eq 0){
						$tempdestinations += $tempdestination
						if ($srule.policy -eq "allow"){
							$tempruleobj = [pscustomobject] $ruleexceptionsobj
							$tempruleobj.priority = $c
							$tempruleobj.dstip = $tempdestination.subnetip
							$tempruleobj.dstsubnet = $tempdestination.subnetmask
							$tempruleobj.direction = 'out'
							$tempruleobj.name = $srule.comment
							$tempruleobj.action = $srule.policy
							$tempruleobj.destination = $dst
							$allowlist += $tempruleobj
						}elseif($srule.policy -eq "deny"){
							$tempruleobj = [pscustomobject] $ruleexceptionsobj
							$tempruleobj.priority = $c
							$tempruleobj.dstip = $tempdestination.subnetip
							$tempruleobj.dstsubnet = $tempdestination.subnetmask
							$tempruleobj.direction = 'out'
							$tempruleobj.name = $srule.comment
							$tempruleobj.action = $srule.policy
							$tempruleobj.destination = $dst
							$denylist += $tempruleobj
						}
					}
				}
				$cs+=1
			}
			
			foreach ($allowrule in $allowlist){
				$tempssidobj.arules += $allowrule
			}
			foreach ($denyrule in $denylist){
				$tempssidobj.drules += $denyrule
			}
			$ssidlist += $tempssidobj
		}
	}
	
	foreach ($s in $ssidlist){
		$s | select name,subnet,networktype,vlan,isguest | ft
		foreach ($s2 in $subnets){
			$rulecompareallow = ''
			$rulecomparedeny = ''
			$sparts = $s2.subnet.split('/')
			$tempip = [ipaddress]$sparts[0]
			$tempsubnet = ConvertTo-IPv4MaskString($sparts[1])

			foreach ($ar in $s.arules){

				if ($ar.dstip.address -eq ($tempip.address -band $ar.dstsubnet.address)){
					if ($ar.priority -lt $rulecompareallow.priority){
						$rulecompareallow = $ar
					}elseif($rulecompareallow -eq ''){
						$rulecompareallow = $ar
					}
				}
			}
			
			foreach ($ad in $s.drules){

				if ($ad.dstip.address -eq ($tempip.address -band $ad.dstsubnet.address)){
					if ($ad.priority -lt $rulecompareallow.priority){
						$rulecomparedeny = $ad
					}elseif($rulecomparedeny = ''){
						$rulecomparedeny = $ad
					}
				}
			}
			if ($rulecompareallow.priority -lt $rulecomparedeny.priority){
				write-output ("      |_ Has access to " + $s2.subnet + " [" + $s2.Name + "] through rule " + $rulecompareallow.name)
			}elseif (($rulecompareallow -ne '') -and ($rulecomparedeny -eq '')){
				write-output ("      |_ Has access to " + $s2.subnet + " [" + $s2.Name + "] through rule " + $rulecompareallow.name)
			}

		
		}
			
	}
	Write-output("================================================================================")
	write-output("")
	Write-output("================================================================================")
	write-output("GROUP POLICIES")
	write-output("Traffic is restricted by most Meraki devices per these policies")
	Write-output('--------------------------------------------------------------------------------')
	$clientlist = @()
	$outputobj = [ordered]@{
		Name = ''
		Subnet = ''
		RuleComment = ''
		GPOName = ''
	}
		
	#Group POLICIES
	$gpo = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/networks/' + $n.NetworkID + '/groupPolicies')
	$gpo = $gpo | convertfrom-json
	$clientpolicies = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/networks/' + $n.NetworkID + '/policies/byClient')
	$clientpolicies = $clientpolicies.content | convertfrom-json
	foreach ($c in $clientpolicies){
		$client = invoke-webrequest -headers $headers -method GET ('https://api.meraki.com/api/v1/networks/' + $n.networkID + '/clients/' + $c.clientID)
		$client = $client.content | convertfrom-json
		$allowed = @()
		$denied = @()
		write-output($c.Name + " " + $client.IP + " " + $c.mac)
		foreach ($pol in $c.assigned){
			$poldeets = $gpo | where {$_.groupPolicyId -eq $pol.groupPolicyID}
			$rules = $poldeets.firewallandtrafficshaping.l3firewallrules
			foreach ($rule in $rules){
				$ipparts = $rule.destCidr.split('/')
				$dstip = [ipaddress] $ipparts[0]
				$dstsub = ConvertTo-IPv4MaskString($ipparts[1])
				
				
				foreach ($s in $subnets){
					$tempipparts = $s.subnet.split('/')
					$tempip = [ipaddress] $tempipparts[0]
					$tempsubnet = ConvertTo-IPv4MaskString($tempipparts[1])
					
					if ($tempip.address -eq ($dstip.address -band $tempsubnet.address)){
						$tempoutputobj = [pscustomobject]$outputobj
						$tempoutputobj.name = $s.Name
						$tempoutputobj.subnet = $s.Subnet
						$tempoutputobj.RuleComment = $rule.comment
						$tempoutputobj.GPOName = $c.name
						if ($rule.policy -eq 'allow'){
							$allowed += $tempoutputobj
						}else{
							$denied += $tempoutputobj
						}
					}
				}
			}
		}
		foreach ($allowrule in $allowed){
			write-output("      |_ Has access to " + $allowrule.subnet + " (" + $allowrule.name + ') through ACL Rule "' + $tempoutputobj.RuleComment + '" in GPO "' + $tempoutputobj.GPOName + '"')
		}
		foreach ($allowrule in $denied){
			write-output("      |_ Is denied access to " + $allowrule.subnet + " (" + $allowrule.name + ') through ACL Rule "' + $tempoutputobj.RuleComment + '" in GPO "' + $tempoutputobj.GPOName + '"')
		}
				
	}
	
}

	
	

	