function Use-Slmgr {
    [cmdletBinding()]
    PARAM (
        [Parameter()] $isKms = $false,
        [Parameter(Mandatory=$true)] $productKey,
        # TODO: Use & nslookup -type=SRV _VLMCS._TCP
        # PSv3: Use Resolve-DnsName ("_VLMCS._TCP." + (Get-DnsClientGlobalSetting).SuffixSearchList) -Type SRV
        [Parameter()] $kmsServer = ""
    )
    
    PROCESS
    {
        # Uninstall Microsoft Product Key
        cscript $env:WINDIR\system32\slmgr.vbs /upk

        # Install Microsoft Product Key
        cscript $env:WINDIR\system32\slmgr.vbs /ipk $productKey

        if ($isKms)
        {
            # Specify the KMS Server
            cscript $env:WINDIR\system32\slmgr.vbs /skms $kmsServer
        }

        # Activate Microsoft Product Key
        cscript $env:WINDIR\system32\slmgr.vbs /ato
    }            
}

# TODO: Create a hash table for MAK/KMS key centralization
# key:value || caption:serial
function Install-WindowsProductKey {
    [CmdletBinding()]
    PARAM (
        [Parameter()] $version = (Get-WmiObject -Class Win32_OperatingSystem).Version,
        [Parameter()] $caption = (Get-WmiObject -Class Win32_OperatingSystem).Caption
    )

    PROCESS
    {
        Write-Verbose -Message "Utilizing $caption as version $version"

        switch -Wildcard ($version)
        {
            # Source (5/11/2015): https://msdn.microsoft.com/en-us/library/windows/desktop/ms724832%28v=vs.85%29.aspx

            "10.0.*"
            {
                Write-Verbose -Message "Windows 10 Insider Preview or Windows Server Technical Preview"
                Write-Host -ForegroundColor Yellow "The Operating System version ($version) is not production."
            }
            
            "6.3.*"
            {
                Write-Verbose -Message "Windows 8.1 or Windows Server 2012 R2"
                if ($caption -like "*Windows Server 2012 R2 Standard*")
                {
                    Write-Verbose -Message "Installing Windows Server 2012 R2 Standard product key."
                    Use-Slmgr -productKey ""
                }
                elseif ($caption -like "*Windows Server 2012 R2 Data*")
                {
                    Write-Verbose -Message "Installing Windows Server 2012 R2 Datacenter product key."
                    Use-Slmgr -productKey ""
                }
                elseif ($caption -like "*Windows 8.1*")
                {
                    Write-Verbose -Message "Installing Windows 8.1 product key."
                    Use-Slmgr -productKey ""
                }
            }

            "6.2.*"
            {
                Write-Verbose -Message "Windows 8 or Windows Server 2012"
                if ($caption -like "*Windows Server 2012*")
                {
                    Write-Verbose -Message "Installing Windows Server 2012 product key."
                    Use-Slmgr -productKey ""
                }
                elseif ($caption -like "*Windows 8*")
                {
                    Write-Verbose -Message "Installing Windows 8 product key."
                    Use-Slmgr -productKey ""
                }
            }

            "6.1.*"
            {
                Write-Verbose -Message "Windows 7 or Windows Server 2008 R2"
                if ($caption -like "*Windows Server 2008 R2*")
                {
                    Write-Verbose -Message "Installing Windows Server 2008 R2 product key."
                    Use-Slmgr -productKey ""
                }
                elseif ($caption -like "*Windows 7*")
                {
                    Write-Verbose -Message "Installing Windows 7 KMS product key and registering to KMS server."
                    Use-Slmgr -isKms $true -productKey ""
                }
            }

            "6.0.*"
            {
                Write-Verbose -Message "Windows Vista or Windows Server 2008"
                if ($caption -like "*Windows Server 2008*")
                {
                    Write-Verbose -Message "Installing Windows Server 2008 product key."
                    Use-Slmgr -productKey ""
                }
                elseif ($caption -like "*Windows Vista*")
                {
                    Write-Host -ForegroundColor Red "The Operating System version ($version) is not supported."
                }
            }

            "5.2.*"
            {
                Write-Verbose -Message "Windows XP 64-Bit Edition, Windows Server 2003, or Windows Server 2003 R2"
                Write-Host -ForegroundColor Red "The Operating System version ($version) is End-of-Life and not supported."
            }

            "5.1.*"
            {
                Write-Verbose -Message "Windows XP"
                Write-Host -ForegroundColor Red "The Operating System version ($version) is End-of-Life and not supported."
            }

            "5.0.*"
            {
                Write-Verbose -Message "Windows 2000"
                Write-Host -ForegroundColor Red "The Operating System version ($version) is End-of-Life and not supported."
            }

            default
            {
                Write-Host -ForegroundColor Red "The Operating System version ($version) is not currently supported."
            }
        }
    }
}
