# 设置控制台编码为 UTF-8
$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

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
            }
        } else {
            Write-Warning "未找到活动的 OpenVPN TAP-Windows 网卡。"
        }
    } catch {
        Write-Warning "获取 IP 地址时发生错误：$($_.Exception.Message)"
    }
    return $null
}

# 获取并过滤 WSL 的有效 IP 地址（仅获取 eth0 接口的 IP）
function Get-WSLIPAddress {
    try {
        $wsl_ip_output = wsl ip -4 addr show eth0 2>&1 | Out-String
        if ($wsl_ip_output -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}") {
            return $matches[0].Trim()
        } else {
            Write-Warning "未能找到 WSL 的有效 IP 地址。"
        }
    } catch {
        Write-Warning "获取 WSL IP 地址时发生错误：$($_.Exception.Message)"
    }
    return $null
}

# 批量执行需要管理员权限的命令（使用单次 sudo 执行所有命令）
function Execute-AdminCommands {
    param (
        [string[]]$Commands
    )

    # 将所有命令放在一个 Here-String 中
    $script = @"
`$ErrorActionPreference = 'Stop'
try {
    $($Commands -join "`r`n")
} catch {
    Write-Error "`$($_.Exception.Message)"
}
"@

    # 使用 sudo 执行所有命令，并将输出重定向到 $null 以抑制不必要的输出
    sudo powershell -Command $script | Out-Null
}

# 添加端口映射和防火墙规则
function Add-PortMapping {
    param (
        [Parameter(Mandatory)]
        [int]$Port
    )

    $wsl_ip = Get-WSLIPAddress
    if ($wsl_ip) {
        $commands = @(
            "netsh interface portproxy delete v4tov4 listenport=$Port listenaddress=0.0.0.0 2>&1 | Out-Null",
            "netsh interface portproxy add v4tov4 listenport=$Port listenaddress=0.0.0.0 connectport=$Port connectaddress=$wsl_ip 2>&1 | Out-Null",
            "if (Get-NetFirewallRule -DisplayName 'WSL $Port' -ErrorAction SilentlyContinue) { Remove-NetFirewallRule -DisplayName 'WSL $Port' -ErrorAction SilentlyContinue }",
            "New-NetFirewallRule -DisplayName 'WSL $Port' -Direction Inbound -Action Allow -Protocol TCP -LocalPort $Port -ErrorAction SilentlyContinue"
        )

        # 批量执行所有命令
        Execute-AdminCommands -Commands $commands
        
        Write-Host "端口 $Port 已从所有接口 (0.0.0.0) 映射到 ${wsl_ip}:${Port}。" -ForegroundColor Green
        Write-Host "已创建防火墙规则，允许端口 $Port 的流量。" -ForegroundColor Green
    } else {
        Write-Warning "无法添加端口映射，因为未找到有效的 WSL IP。"
    }
}

# 删除端口映射和防火墙规则
function Remove-PortMapping {
    param (
        [Parameter(Mandatory)]
        [int]$Port,
        [switch]$Quiet
    )

    $mappings = netsh interface portproxy show all
    $relatedMappings = $mappings | ForEach-Object {
        if ($_ -match "(\S+)\s+$Port\s+\S+\s+$Port") {
            $matches[1]
        }
    } | Sort-Object -Unique

    $commands = @()

    if ($relatedMappings) {
        foreach ($listenAddress in $relatedMappings) {
            $commands += "netsh interface portproxy delete v4tov4 listenport=$Port listenaddress=$listenAddress 2>&1 | Out-Null"
        }
    }

    # 检查防火墙规则是否存在，然后删除
    $commands += "if (Get-NetFirewallRule -DisplayName 'WSL $Port' -ErrorAction SilentlyContinue) { Remove-NetFirewallRule -DisplayName 'WSL $Port' -ErrorAction SilentlyContinue }"

    # 批量执行所有命令
    Execute-AdminCommands -Commands $commands
    
    Write-Host "与端口 ${Port} 相关的映射和防火墙规则已删除。" -ForegroundColor Green
}

# 列出所有符合 "WSL + PORT" 规则的防火墙规则
function List-WSLFirewallRules {
    try {
        $rules = Get-NetFirewallRule | Where-Object { $_.DisplayName -like 'WSL *' }
        if ($rules) {
            Write-Host "当前符合 'WSL + PORT' 规则的防火墙规则：" -ForegroundColor Green
            $rules | Format-Table -Property DisplayName, Direction, Action, Enabled, Profile -AutoSize
        } else {
            Write-Host "未找到符合 'WSL + PORT' 规则的防火墙规则。" -ForegroundColor Yellow
        }
    } catch {
        Write-Warning "列出防火墙规则时发生错误：$($_.Exception.Message)"
    }
}

# 根据端口号删除防火墙规则
function Remove-FirewallRuleByPort {
    param (
        [Parameter(Mandatory)]
        [int]$Port
    )

    try {
        $ruleName = "WSL $Port"
        if (Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue) {
            Execute-AdminCommands -Commands @("Remove-NetFirewallRule -DisplayName '$ruleName' -ErrorAction SilentlyContinue")
            Write-Host "防火墙规则 '$ruleName' 已删除。" -ForegroundColor Green
        } else {
            Write-Host "未找到防火墙规则 '$ruleName'。" -ForegroundColor Yellow
        }
    } catch {
        Write-Warning "删除防火墙规则时发生错误：$($_.Exception.Message)"
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

        $tap_ip = Get-TapIpAddress
        if ($tap_ip) {
            Write-Host "OpenVPN TAP-Windows IP 地址: $tap_ip" -ForegroundColor Cyan
        } else {
            Write-Warning "未能找到 OpenVPN TAP-Windows IP 地址。"
        }

        $wsl_ip = Get-WSLIPAddress
        if ($wsl_ip) {
            Write-Host "WSL IP 地址: $wsl_ip" -ForegroundColor Cyan
        } else {
            Write-Warning "未能找到 WSL IP 地址。"
        }

        Write-Host "====================================="
        Write-Host "请选择操作："
        Write-Host "1. 添加端口映射"
        Write-Host "2. 删除端口映射"
        Write-Host "3. 列出所有端口映射"
        Write-Host "4. 列出所有 WSL 防火墙规则"
        Write-Host "5. 根据端口删除防火墙规则"
        Write-Host "6. 退出"
        Write-Host "====================================="
        $choice = Read-Host "请输入你的选择 (1-6)"

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
            "4" { List-WSLFirewallRules }
            "5" {
                $port = Read-Host "请输入要删除的防火墙规则的端口"
                Remove-FirewallRuleByPort -Port $port
            }
            "6" {
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
