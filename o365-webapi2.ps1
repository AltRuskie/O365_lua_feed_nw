# webservice root URL
$ws = "https://endpoints.office.com"

# name of the feed and parser details for pivots
$rsa_name = "office365"

# filename parameters
$Date = Get-Date
$DateString = $Date.ToString("hhmmdd-mm-yyyy")
$Filepath = "$PSScriptRoot\"

# path where client ID and latest version number will be stored
$datapath = $Env:TEMP + "\endpoints_clientid_latestversion.txt"
Write-Output "Flags data written to this path "$datapath

# fetch client ID and version if data file exists; otherwise create new file
if (Test-Path $datapath) {
    $content = Get-Content $datapath
    $clientRequestId = $content[0]
    $lastVersion = $content[1]
}
else {
    $clientRequestId = [GUID]::NewGuid().Guid
    $lastVersion = "0000000000"
    @($clientRequestId, $lastVersion) | Out-File $datapath
}

# call version method to check the latest version, and pull new data if version number is different
$version = Invoke-RestMethod -Uri ($ws + "/version/Worldwide?clientRequestId=" + $clientRequestId)
if ($version.latest -gt $lastVersion) {
    Write-Host "New version of Office 365 worldwide commercial service instance endpoints detected"
    
    # write the new version number to the data file
    @($clientRequestId, $version.latest) | Out-File $datapath

    # invoke endpoints method to get the new data
    $endpointSets = Invoke-RestMethod -Uri ($ws + "/endpoints/Worldwide?clientRequestId=" + $clientRequestId)

	Write-Host "RAW Output"
    $endpointSets | Out-String
	
    # filter results for Allow and Optimize endpoints, and transform these into custom objects with port and category
    $flatUrls = $endpointSets | ForEach-Object {
        $endpointSet = $_
        $urls = $(if ($endpointSet.urls.Count -gt 0) { $endpointSet.urls } else { @() })

        $urlCustomObjects = @()
        if ($endpointSet.category -in ("Allow", "Optimize", "Default")) {
            $urlCustomObjects = $urls | ForEach-Object {
                [PSCustomObject]@{
                    category = $endpointSet.category;
                    url      = $_;
                    #tcpPorts = $endpointSet.tcpPorts;
                    #udpPorts = $endpointSet.udpPorts;
					# added the service area
					serviceArea = $endpointSet.serviceArea;
					# serviceAreaDispayName = $endpointSet.serviceAreaDispayName;
                }
            }
        }
        $urlCustomObjects
    }

    $flatIps = $endpointSets | ForEach-Object {
        $endpointSet = $_
        $ips = $(if ($endpointSet.ips.Count -gt 0) { $endpointSet.ips } else { @() })
        # IPv4 strings have dots while IPv6 strings have colons
		
		# select IPv4 addresses
        $ip4s = $ips | Where-Object { $_ -like '*.*' }
		
        $ipCustomObjects = @()
        if ($endpointSet.category -in ("Allow", "Optimize", "Default")) {
            $ipCustomObjects = $ip4s | ForEach-Object {
                [PSCustomObject]@{
                    category = $endpointSet.category;
                    ip = $_;
                    #tcpPorts = $endpointSet.tcpPorts;
                    #udpPorts = $endpointSet.udpPorts;
					# added the service area
					serviceArea = $endpointSet.serviceArea;
					# serviceAreaDispayName = $endpointSet.serviceAreaDispayName;
                }
            }
        }
        $ipCustomObjects
	}
	
	$flatIps6 = $endpointSets | ForEach-Object {
        $endpointSet = $_
        $ips6 = $(if ($endpointSet.ips6.Count -gt 0) { $endpointSet.ips6 } else { @() })
        # IPv4 strings have dots while IPv6 strings have colons
		# select IPv6 addresses
        $ip6s = $ips6 | Where-Object { $_ -like '*:*' }
		
        $ipCustomObjects6 = @()
        if ($endpointSet.category -in ("Allow", "Optimize", "Default")) {
            $ipCustomObjects6 = $ip6s | ForEach-Object {
                [PSCustomObject]@{
                    category = $endpointSet.category;
                    ip = $_;
                    #tcpPorts = $endpointSet.tcpPorts;
                    #udpPorts = $endpointSet.udpPorts;
					# added the service area
					serviceArea = $endpointSet.serviceArea;
					# serviceAreaDispayName = $endpointSet.serviceAreaDispayName;
                }
            }
        }
        $ipCustomObjects6
    }

    Write-Output "IPv4 Firewall IP Address Ranges"
    #($flatIps.ip | Sort-Object -Unique) -join "," | Out-String
    #($flatIps.ip | Sort-Object -Unique) -join ",$rsa_name`n" | Out-String
    ($flatIps.ip | Sort-Object -Unique) -join ",whitelist,$rsa_name`n" | Tee-Object -FilePath $Filepath"o365ipv4Out.csv"
	
	Write-Output "IPv6 Firewall IP Address Ranges"
    #($flatIps6.ip | Sort-Object -Unique) -join "," | Out-String
    #($flatIps6.ip | Sort-Object -Unique) -join ",$rsa_name`n" | Out-String
    ($flatIps6.ip | Sort-Object -Unique) -join ",whitelist,$rsa_name`n" | Tee-Object -FilePath $Filepath"o365ipv6Out.csv"

    Write-Output "URLs for Proxy Server"
    #($flatUrls.url | Sort-Object -Unique) -join "," | Out-String
    #($flatUrls.url | Sort-Object -Unique) -join "`"] = `"$rsa_name`",`n[`"" | Out-String
    #($flatUrls.url | Sort-Object -Unique) -join "`"] = `"$rsa_name`",`n[`"" | Tee-Object -FilePath $Filepath"o365urlOut.txt"
    $urlVar = ($flatUrls.url | Sort-Object -Unique) -join "`"] = `"$rsa_name`",`n[`"" | Out-String
	# replace wildcards with nothing
	$urlVar = $urlVar -replace "\*\.", ""
	# prepend and append to the string to format for lua parser
    ("[`"{0}`"] = `"office365`"" -f $urlVar) | Tee-Object -FilePath $Filepath"o365urlOut.txt"


    # TODO Call Send-MailMessage with new endpoints data
}
else {
    Write-Host "Office 365 worldwide commercial service instance endpoints are up-to-date"
}