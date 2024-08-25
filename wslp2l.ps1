# 设置控制台编码为 UTF-8
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 获取 OpenVPN TAP-Windows 网卡的 IP 地址
function Get-TapIpAddress {
    try {
        $adapter = Get-NetAdapter | Where-Object { $_.Name -match "TAP-Windows" -and $_.Status -eq "Up" }
        if ($adapter) {
            $ipAddress = (Get-NetIPAddress -InterfaceIndex $adapter.InterfaceIndex -AddressFamily IPv4).IPAddress
            if ($ipAddress -match "^192\.168\..*") {
                return $ipAddress
            } else {
                Write-Warning "未找到符合条件的 OpenVPN TAP-Windows 网卡的 IPv4 地址。"
                return $null
            }
        } else {
            Write-Warning "未找到活动的 OpenVPN TAP-Windows 网卡。"
            return $null
        }
    } catch {
        Write-Warning "获取 IP 地址时发生错误：$($_.Exception.Message)"
        return $null
    }
}

# 添加端口映射
function Add-PortMapping {
    param (
        [Parameter(Mandatory)]
        [int]$Port
    )
    
    $tap_ip = Get-TapIpAddress
    if ($tap_ip) {
        $cmd = "netsh interface portproxy add v4tov4 listenport=$Port listenaddress=$tap_ip connectport=$Port connectaddress=localhost"
        try {
            Start-Process -FilePath "powershell" -ArgumentList "-Command `"$cmd`"" -Verb RunAs
            Write-Host "端口 $Port 已从 $tap_ip 映射到本地端口 localhost:$Port。" -ForegroundColor Green
        } catch {
            Write-Warning "添加端口映射失败：$($_.Exception.Message)"
        }
    } else {
        Write-Warning "无法添加端口映射，因为未找到有效的 OpenVPN TAP-Windows IP。"
    }
}

# 删除端口映射
function Remove-PortMapping {
    param (
        [Parameter(Mandatory)]
        [int]$Port
    )
    
    # 获取所有映射条目
    $mappings = netsh interface portproxy show all

    # 查找与指定端口相关的所有映射
    $relatedMappings = $mappings | ForEach-Object {
        if ($_ -match "(\S+)\s+$Port\s+\S+\s+$Port") {
            $listenAddress = $matches[1]
            return $listenAddress
        }
    }

    if ($relatedMappings) {
        $relatedMappings = $relatedMappings | Sort-Object -Unique
        foreach ($listenAddress in $relatedMappings) {
            $cmd = "netsh interface portproxy delete v4tov4 listenport=$Port listenaddress=$listenAddress"
            try {
                Start-Process -FilePath "powershell" -ArgumentList "-Command `"$cmd`"" -Verb RunAs
                Write-Host "与 ${listenAddress}:${Port} 相关的映射已删除。" -ForegroundColor Green
            } catch {
                Write-Warning "删除端口映射失败：$($_.Exception.Message)"
            }
        }
    } else {
        Write-Warning "未找到与端口 $Port 相关的映射。"
    }
}

# 列出所有端口映射
function List-PortMappings {
    Write-Host "当前的端口映射：" -ForegroundColor Green
    netsh interface portproxy show all | Format-Table -AutoSize
}


# 主菜单函数
function Show-Menu {
    while ($true) {
        Clear-Host
        
        # 每次显示菜单前输出OpenVPN IP
        $tap_ip = Get-TapIpAddress
        if ($tap_ip) {
            Write-Host "OpenVPN TAP-Windows IP 地址: $tap_ip" -ForegroundColor Cyan
        } else {
            Write-Warning "未能找到 OpenVPN TAP-Windows IP 地址。"
        }
        
        Write-Host "====================================="
        Write-Host "请选择操作："
        Write-Host "1. 添加端口映射"
        Write-Host "2. 删除端口映射"
        Write-Host "3. 列出所有端口映射"
        Write-Host "4. 退出"
        Write-Host "====================================="
        $choice = Read-Host "请输入你的选择 (1-4)"

        switch ($choice) {
            "1" {
                $port = Read-Host "请输入要添加的端口"
                Add-PortMapping -Port $port
            }
            "2" {
                $port = Read-Host "请输入要删除的端口"
                Remove-PortMapping -Port $port
            }
            "3" { List-PortMappings }
            "4" {
                Write-Host "正在退出..." -ForegroundColor Yellow
                Start-Sleep -Seconds 1
                exit
            }
            default { Write-Host "无效的选择，请重新输入。" -ForegroundColor Red }
        }
        Pause
    }
}

# 执行主菜单
Show-Menu
