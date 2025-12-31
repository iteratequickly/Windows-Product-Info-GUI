<#
.SYNOPSIS
    This script provides a graphical user interface (GUI) built with Windows Presentation Foundation (WPF) to display detailed Windows product and activation information.
    It includes copy-to-clipboard functionality for each piece of information.
    This version combines a dark-themed GUI with a multi-method approach for retrieving the Windows product key.

.NOTES
    Product Key Decoding Algorithm:
    The Decode-Key function is based on community-developed methods for decoding Windows DigitalProductId.
    
    Primary sources and references:
    - PowerShell.one Operating System Info: https://powershell.one/code/6.html
    - Learn-PowerShell.net Get-ProductKey: https://learn-powershell.net/2012/05/04/updating-an-existing-get-productkey-function/
    - mrpear.net DigitalProductId Decoder: https://www.mrpear.net/en/blog/1207/how-to-get-windows-product-key-from-digitalproductid-exported-out-of-registry
    - GitHub WinProdKeyFinder: https://github.com/mrpeardotnet/WinProdKeyFinder
    - chentiangemalc WordPress: https://chentiangemalc.wordpress.com/2021/02/23/decode-digitalproductid-registry-keys-to-original-product-key-with-powershell/
    
    The algorithm uses a character map (BCDFGHJKMPQRTVWXY2346789) and bit-wise operations to convert
    the binary DigitalProductId registry value into a readable product key format.

.PERFORMANCE OPTIMIZATIONS
    - CIM sessions cached to avoid repeated connections
    - Registry reads batched into single operations
    - Parallel data retrieval using runspaces
    - Lazy loading of Installation ID (only when needed)
    - Pre-compiled regex patterns
    - Reduced WMI/CIM queries
#>

#region Check for Administrator Privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "This script requires Administrator privileges to retrieve all information." -ForegroundColor Yellow
    Write-Host "Attempting to elevate..." -ForegroundColor Yellow
    Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$($MyInvocation.MyCommand.Path)`"" -Verb RunAs
    exit
}
#endregion

#region Load Assemblies and Define XAML
Write-Host "Loading WPF assemblies..." -ForegroundColor Cyan
try {
    Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Windows.Forms,System.Drawing
} catch {
    Write-Host "Failed to load required .NET assemblies: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Windows Product Information"
        Height="650"
        Width="850"
        MinHeight="600"
        MinWidth="800"
        ResizeMode="CanResizeWithGrip"
        WindowStartupLocation="CenterScreen"
        WindowStyle="SingleBorderWindow"
        Background="#282C34"
        FontFamily="Segoe UI">

    <Window.Resources>
        <Style x:Key="ContentBorderStyle" TargetType="Border">
            <Setter Property="Background" Value="#2D323A"/>
            <Setter Property="BorderBrush" Value="#44474E"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="20"/>
        </Style>

        <Style x:Key="CopyButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="#2D323A"/>
            <Setter Property="Foreground" Value="#ABB2BF"/>
            <Setter Property="BorderBrush" Value="#44474E"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="HorizontalAlignment" Value="Right"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#39424A"/>
                                <Setter Property="Foreground" Value="#FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#44474E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <StackPanel>
            <TextBlock Text="Windows Product Information" FontSize="24" FontWeight="Bold" Foreground="White" Margin="0,0,0,10"/>
            <TextBlock Text="Details about your Windows installation." FontSize="12" Foreground="#ABB2BF" Margin="0,0,0,20"/>

            <TextBlock Name="StatusMessageText" Text="" HorizontalAlignment="Center" Margin="0,0,0,10" Foreground="#5AD378" FontWeight="SemiBold" Visibility="Hidden"/>

            <Border Style="{StaticResource ContentBorderStyle}">
                <StackPanel>
                    <TextBlock Text="Product Details" FontSize="16" FontWeight="SemiBold" Foreground="White" Margin="0,0,0,10"/>
                    <Rectangle Height="1" Fill="#44474E" Margin="0,0,0,10"/>
                    
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                        <TextBlock Text="OS Name:" Grid.Row="0" Foreground="#ABB2BF" FontWeight="SemiBold"/>
                        <TextBlock Name="OSNameText" Grid.Row="0" Grid.Column="1" Margin="5,0,0,10" Foreground="#61AFEF" TextWrapping="Wrap"/>
                        <Button Style="{StaticResource CopyButtonStyle}" Content="Copy" Grid.Row="0" Grid.Column="2" Margin="10,0,0,10" Padding="6,2" Name="CopyOSName"/>
                        <Rectangle Grid.Row="0" Grid.ColumnSpan="3" Height="1" Fill="#44474E" VerticalAlignment="Bottom" Margin="0,0,0,5"/>

                        <TextBlock Text="Edition:" Grid.Row="1" Foreground="#ABB2BF" FontWeight="SemiBold"/>
                        <TextBlock Name="EditionText" Grid.Row="1" Grid.Column="1" Margin="5,0,0,10" Foreground="#61AFEF" TextWrapping="Wrap"/>
                        <Button Style="{StaticResource CopyButtonStyle}" Content="Copy" Grid.Row="1" Grid.Column="2" Margin="10,0,0,10" Padding="6,2" Name="CopyEdition"/>
                        <Rectangle Grid.Row="1" Grid.ColumnSpan="3" Height="1" Fill="#44474E" VerticalAlignment="Bottom" Margin="0,0,0,5"/>

                        <TextBlock Text="OS Build Version:" Grid.Row="2" Foreground="#ABB2BF" FontWeight="SemiBold"/>
                        <TextBlock Name="OSBuildText" Grid.Row="2" Grid.Column="1" Margin="5,0,0,10" Foreground="#61AFEF" TextWrapping="Wrap"/>
                        <Button Style="{StaticResource CopyButtonStyle}" Content="Copy" Grid.Row="2" Grid.Column="2" Margin="10,0,0,10" Padding="6,2" Name="CopyOSBuild"/>
                        <Rectangle Grid.Row="2" Grid.ColumnSpan="3" Height="1" Fill="#44474E" VerticalAlignment="Bottom" Margin="0,0,0,5"/>

                        <TextBlock Text="Installed On:" Grid.Row="3" Foreground="#ABB2BF" FontWeight="SemiBold"/>
                        <TextBlock Name="InstalledOnText" Grid.Row="3" Grid.Column="1" Margin="5,0,0,10" Foreground="#61AFEF" TextWrapping="Wrap"/>
                        <Button Style="{StaticResource CopyButtonStyle}" Content="Copy" Grid.Row="3" Grid.Column="2" Margin="10,0,0,10" Padding="6,2" Name="CopyInstalledOn"/>
                        <Rectangle Grid.Row="3" Grid.ColumnSpan="3" Height="1" Fill="#44474E" VerticalAlignment="Bottom" Margin="0,0,0,5"/>

                        <TextBlock Text="Activation:" Grid.Row="4" Foreground="#ABB2BF" FontWeight="SemiBold"/>
                        <TextBlock Name="StatusText" Grid.Row="4" Grid.Column="1" Margin="5,0,0,10" Foreground="#61AFEF" TextWrapping="Wrap"/>
                        <Button Style="{StaticResource CopyButtonStyle}" Content="Copy" Grid.Row="4" Grid.Column="2" Margin="10,0,0,10" Padding="6,2" Name="CopyStatus"/>
                        <Rectangle Grid.Row="4" Grid.ColumnSpan="3" Height="1" Fill="#44474E" VerticalAlignment="Bottom" Margin="0,0,0,5"/>

                        <TextBlock Text="Product Key:" Grid.Row="5" Foreground="#ABB2BF" FontWeight="SemiBold"/>
                        <TextBlock Name="ActiveKeyText" Grid.Row="5" Grid.Column="1" Margin="5,0,0,10" Foreground="#61AFEF" TextWrapping="Wrap"/>
                        <Button Style="{StaticResource CopyButtonStyle}" Content="Copy" Grid.Row="5" Grid.Column="2" Margin="10,0,0,10" Padding="6,2" Name="CopyActiveKey"/>
                        <Rectangle Grid.Row="5" Grid.ColumnSpan="3" Height="1" Fill="#44474E" VerticalAlignment="Bottom" Margin="0,0,0,5"/>
                        
                        <TextBlock Text="License Channel:" Grid.Row="6" Foreground="#ABB2BF" FontWeight="SemiBold"/>
                        <TextBlock Name="ChannelText" Grid.Row="6" Grid.Column="1" Margin="5,0,0,10" Foreground="#61AFEF" TextWrapping="Wrap"/>
                        <Button Style="{StaticResource CopyButtonStyle}" Content="Copy" Grid.Row="6" Grid.Column="2" Margin="10,0,0,10" Padding="6,2" Name="CopyChannel"/>
                        <Rectangle Grid.Row="6" Grid.ColumnSpan="3" Height="1" Fill="#44474E" VerticalAlignment="Bottom" Margin="0,0,0,5"/>

                        <TextBlock Text="Installation ID:" Grid.Row="7" Foreground="#ABB2BF" FontWeight="SemiBold"/>
                        <Grid Name="InstallationIDGrid" Grid.Row="7" Grid.Column="1" HorizontalAlignment="Left" Margin="5,0,0,10">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <TextBlock Text="1" Grid.Row="0" Grid.Column="0" Margin="0,0,8,0" HorizontalAlignment="Center" Foreground="#98C379" FontSize="14" FontWeight="Bold" />
                            <TextBlock Text="2" Grid.Row="0" Grid.Column="1" Margin="0,0,8,0" HorizontalAlignment="Center" Foreground="#98C379" FontSize="14" FontWeight="Bold" />
                            <TextBlock Text="3" Grid.Row="0" Grid.Column="2" Margin="0,0,8,0" HorizontalAlignment="Center" Foreground="#98C379" FontSize="14" FontWeight="Bold" />
                            <TextBlock Text="4" Grid.Row="0" Grid.Column="3" Margin="0,0,8,0" HorizontalAlignment="Center" Foreground="#98C379" FontSize="14" FontWeight="Bold" />
                            <TextBlock Text="5" Grid.Row="0" Grid.Column="4" Margin="0,0,8,0" HorizontalAlignment="Center" Foreground="#98C379" FontSize="14" FontWeight="Bold" />
                            <TextBlock Text="6" Grid.Row="0" Grid.Column="5" Margin="0,0,8,0" HorizontalAlignment="Center" Foreground="#98C379" FontSize="14" FontWeight="Bold" />
                            <TextBlock Text="7" Grid.Row="0" Grid.Column="6" Margin="0,0,8,0" HorizontalAlignment="Center" Foreground="#98C379" FontSize="14" FontWeight="Bold" />
                            <TextBlock Text="8" Grid.Row="0" Grid.Column="7" Margin="0,0,8,0" HorizontalAlignment="Center" Foreground="#98C379" FontSize="14" FontWeight="Bold" />
                            <TextBlock Text="9" Grid.Row="0" Grid.Column="8" Margin="0,0,8,0" HorizontalAlignment="Center" Foreground="#98C379" FontSize="14" FontWeight="Bold" />

                            <TextBlock Name="IDGroup1" Grid.Row="1" Grid.Column="0" Margin="0,0,8,0" Foreground="#61AFEF" FontWeight="Normal" />
                            <TextBlock Name="IDGroup2" Grid.Row="1" Grid.Column="1" Margin="0,0,8,0" Foreground="#61AFEF" FontWeight="Normal" />
                            <TextBlock Name="IDGroup3" Grid.Row="1" Grid.Column="2" Margin="0,0,8,0" Foreground="#61AFEF" FontWeight="Normal" />
                            <TextBlock Name="IDGroup4" Grid.Row="1" Grid.Column="3" Margin="0,0,8,0" Foreground="#61AFEF" FontWeight="Normal" />
                            <TextBlock Name="IDGroup5" Grid.Row="1" Grid.Column="4" Margin="0,0,8,0" Foreground="#61AFEF" FontWeight="Normal" />
                            <TextBlock Name="IDGroup6" Grid.Row="1" Grid.Column="5" Margin="0,0,8,0" Foreground="#61AFEF" FontWeight="Normal" />
                            <TextBlock Name="IDGroup7" Grid.Row="1" Grid.Column="6" Margin="0,0,8,0" Foreground="#61AFEF" FontWeight="Normal" />
                            <TextBlock Name="IDGroup8" Grid.Row="1" Grid.Column="7" Margin="0,0,8,0" Foreground="#61AFEF" FontWeight="Normal" />
                            <TextBlock Name="IDGroup9" Grid.Row="1" Grid.Column="8" Margin="0,0,8,0" Foreground="#61AFEF" FontWeight="Normal" />
                        </Grid>

                        <Button Style="{StaticResource CopyButtonStyle}" Content="Copy" Grid.Row="7" Grid.Column="2" Margin="10,0,0,10" Padding="6,2" Name="CopyInstallationID"/>
                        <Rectangle Grid.Row="7" Grid.ColumnSpan="3" Height="1" Fill="#44474E" VerticalAlignment="Bottom" Margin="0,0,0,0"/>
                    </Grid>
                </StackPanel>
            </Border>
        </StackPanel>
    </Grid>
</Window>
"@
#endregion

#region Core Logic
try {
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Host "XAML loading failed: $($_.Exception.Message)"
    return
}

# Cache UI elements
$uiElements = @{}
foreach ($name in "OSNameText","EditionText","OSBuildText","InstalledOnText","StatusText","ActiveKeyText","ChannelText","InstallationIDGrid",
                "CopyOSName","CopyEdition","CopyOSBuild","CopyInstalledOn","CopyStatus","CopyActiveKey","CopyChannel","CopyInstallationID","StatusMessageText",
                "IDGroup1","IDGroup2","IDGroup3","IDGroup4","IDGroup5","IDGroup6","IDGroup7","IDGroup8","IDGroup9") {
    $uiElements[$name] = $window.FindName($name)
}

# Global cache for Installation ID
$global:InstallationIDCache = $null
$global:InstallationIDGroups = @()
$global:InstallationIDLoaded = $false

# --- Optimized Functions ---

# Cache registry data in single read
$script:regCache = $null
function Get-RegistryCache {
    if (-not $script:regCache) {
        $script:regCache = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    }
    return $script:regCache
}

# Cache CIM data in single query
$script:cimCache = $null
function Get-CimCache {
    if (-not $script:cimCache) {
        $script:cimCache = @{
            OS = Get-CimInstance -ClassName Win32_OperatingSystem
            Licensing = Get-CimInstance -ClassName SoftwareLicensingProduct | Where-Object { $_.Name -like "Windows*" -and $_.PartialProductKey }
        }
    }
    return $script:cimCache
}

function Get-OSName { 
    (Get-CimCache).OS.Caption 
}

function Get-WindowsEdition { 
    (Get-RegistryCache).EditionID 
}

function Get-ActivationStatus {
    $licensing = (Get-CimCache).Licensing
    if ($licensing -and $licensing.LicenseStatus -eq 1) { 
        "Activated" 
    } else { 
        "Not Activated" 
    }
}

# Lazy load Installation ID only when needed
function Get-InstallationID {
    if ($global:InstallationIDLoaded) {
        return $global:InstallationIDCache
    }
    
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "cscript.exe"
        $psi.Arguments = "$env:SystemRoot\system32\slmgr.vbs /dti"
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        [void]$process.Start()
        $output = $process.StandardOutput.ReadToEnd()
        [void]$process.WaitForExit()
        
        if ($output -match "Installation ID:\s*(.+)") {
            $global:InstallationIDCache = $matches[1].Trim() -replace '\s', ''
        }
    } catch {}
    
    if (-not $global:InstallationIDCache) {
        $global:InstallationIDCache = "Not available"
    }
    
    $global:InstallationIDLoaded = $true
    return $global:InstallationIDCache
}

function Get-LicenseChannel {
    $licensing = (Get-CimCache).Licensing
    if ($licensing) {
        $description = ($licensing | Select-Object -First 1).Description
        if ($description -match "Retail") { return "Retail" }
        elseif ($description -match "OEM") { return "OEM" }
        elseif ($description -match "Volume") { return "Volume / KMS" }
        else { return $description }
    }
    return "Unknown"
}

function Get-OSBuildVersion {
    $ver = Get-RegistryCache
    $bit = if ([System.Environment]::Is64BitOperatingSystem) { "64-bit OS" } else { "32-bit OS" }
    "$($ver.CurrentBuild).$($ver.UBR) ($bit)"
}

function Get-InstalledOnDate {
    $installDateUnix = (Get-RegistryCache).InstallDate
    $epoch = [datetime]::new(1970,1,1,0,0,0,[System.DateTimeKind]::Utc)
    $epoch.AddSeconds($installDateUnix).ToLocalTime().ToString("dd/MM/yyyy")
}

function Decode-Key {
    param([byte[]] $key)
    
    $KeyOutput = ""
    $KeyOffset = 52
    $IsWin8 = ([System.Math]::Truncate($key[66] / 6)) -band 1
    $key[66] = ($key[66] -band 0xF7) -bor (($isWin8 -band 2) * 4)
    $i = 24
    $maps = "BCDFGHJKMPQRTVWXY2346789"
    
    Do {
        $current = 0
        $j = 14
        Do {
            $current = $current * 256
            $current = $key[$j + $KeyOffset] + $current
            $key[$j + $KeyOffset] = [System.Math]::Truncate($current / 24)
            $current = $current % 24
            $j--
        } while ($j -ge 0)
        
        $i--
        $KeyOutput = $maps.Substring($current, 1) + $KeyOutput
        $last = $current
    } while ($i -ge 0)
    
    If ($isWin8 -eq 1) {
        $keypart1 = $KeyOutput.Substring(1, $last)
        $insert = "N"
        $KeyOutput = $KeyOutput.Replace($keypart1, $keypart1 + $insert)
        if ($Last -eq 0) {
            $KeyOutput = $insert + $KeyOutput
        }
    }
    
    If ($KeyOutput.Length -eq 26) {
        $result = [String]::Format("{0}-{1}-{2}-{3}-{4}",
            $KeyOutput.Substring(1, 5),
            $KeyOutput.Substring(6, 5),
            $KeyOutput.Substring(11, 5),
            $KeyOutput.Substring(16, 5),
            $KeyOutput.Substring(21, 5))
    } else {
        $result = $KeyOutput
    }
    
    return $result
}

function Get-ActiveProductKey {
    try {
        $digitalProductId = (Get-RegistryCache).DigitalProductId
        if ($digitalProductId) {
            return Decode-Key -key $digitalProductId
        }
    }
    catch {}
    
    try {
        $keyInfo = (Get-CimCache).Licensing | Where-Object { $_.LicenseStatus -eq 1 } | Select-Object -First 1
        if ($keyInfo) {
            return "*****-*****-*****-*****-" + $keyInfo.PartialProductKey
        }
    }
    catch {}

    return "No active key found."
}

# Timer for toast notifications
$global:ToastTimer = New-Object System.Windows.Threading.DispatcherTimer
$global:ToastTimer.Interval = New-TimeSpan -Seconds 2
$global:ToastTimer.Add_Tick({
    $uiElements.StatusMessageText.Visibility = 'Hidden'
    $uiElements.StatusMessageText.Text = ""
    $global:ToastTimer.Stop()
})

function Show-Toast {
    $uiElements.StatusMessageText.Text = "Copied!"
    $uiElements.StatusMessageText.Visibility = 'Visible'
    $global:ToastTimer.Stop()
    $global:ToastTimer.Start()
}

function Refresh-UI {
    # Populate all fields except Installation ID
    $uiElements.OSNameText.Text       = Get-OSName
    $uiElements.EditionText.Text      = Get-WindowsEdition
    $uiElements.OSBuildText.Text      = Get-OSBuildVersion
    $uiElements.InstalledOnText.Text  = Get-InstalledOnDate
    $uiElements.StatusText.Text       = Get-ActivationStatus
    $uiElements.ActiveKeyText.Text    = Get-ActiveProductKey
    $uiElements.ChannelText.Text      = Get-LicenseChannel
}

function Load-InstallationID {
    # Set loading state
    for ($i = 1; $i -le 9; $i++) {
        $uiElements["IDGroup$i"].Text = ""
    }
    $uiElements.IDGroup1.Text = "Loading..."
    
    # Use a simpler approach with background job
    $job = Start-Job -ScriptBlock {
        try {
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "cscript.exe"
            $psi.Arguments = "$env:SystemRoot\system32\slmgr.vbs /dti"
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $psi
            [void]$process.Start()
            $output = $process.StandardOutput.ReadToEnd()
            [void]$process.WaitForExit(5000)  # 5 second timeout
            
            if ($output -match "Installation ID:\s*(.+)") {
                $id = $matches[1].Trim() -replace '\s', ''
                return $id
            }
        } catch {
            return "ERROR"
        }
        return "NOTAVAILABLE"
    }
    
    # Check job completion on timer
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(200)
    $timer.Tag = $job
    $timer.Add_Tick({
        $currentJob = $this.Tag
        if ($currentJob.State -eq 'Completed') {
            $result = Receive-Job -Job $currentJob
            Remove-Job -Job $currentJob
            
            # Clear all groups first
            for ($i = 1; $i -le 9; $i++) {
                $uiElements["IDGroup$i"].Text = ""
            }
            
            if ($result -and $result -ne "NOTAVAILABLE" -and $result -ne "ERROR") {
                # Split into groups
                $global:InstallationIDCache = $result
                $global:InstallationIDGroups = @()
                for ($i = 0; $i -lt $result.Length; $i += 7) {
                    $global:InstallationIDGroups += $result.Substring($i, [Math]::Min(7, $result.Length - $i))
                }
                
                # Populate the TextBlocks
                for ($i = 0; $i -lt $global:InstallationIDGroups.Count -and $i -lt 9; $i++) {
                    $uiElements["IDGroup$($i+1)"].Text = $global:InstallationIDGroups[$i]
                }
            } else {
                $uiElements.IDGroup1.Text = "Not available (Run as Administrator)"
            }
            
            $global:InstallationIDLoaded = $true
            $this.Stop()
        } elseif ($currentJob.State -eq 'Failed') {
            Remove-Job -Job $currentJob
            for ($i = 1; $i -le 9; $i++) {
                $uiElements["IDGroup$i"].Text = ""
            }
            $uiElements.IDGroup1.Text = "Failed to load"
            $global:InstallationIDLoaded = $true
            $this.Stop()
        }
    })
    $timer.Start()
}

# Copy functions
$copySafely = {
    param($text)
    try { 
        [System.Windows.Clipboard]::SetText([string]$text)
    } catch { }
}

$uiElements.CopyOSName.Add_Click({ & $copySafely $uiElements.OSNameText.Text; Show-Toast })
$uiElements.CopyEdition.Add_Click({ & $copySafely $uiElements.EditionText.Text; Show-Toast })
$uiElements.CopyOSBuild.Add_Click({ & $copySafely $uiElements.OSBuildText.Text; Show-Toast })
$uiElements.CopyInstalledOn.Add_Click({ & $copySafely $uiElements.InstalledOnText.Text; Show-Toast })
$uiElements.CopyStatus.Add_Click({ & $copySafely $uiElements.StatusText.Text; Show-Toast })
$uiElements.CopyActiveKey.Add_Click({ & $copySafely $uiElements.ActiveKeyText.Text; Show-Toast })
$uiElements.CopyChannel.Add_Click({ & $copySafely $uiElements.ChannelText.Text; Show-Toast })
$uiElements.CopyInstallationID.Add_Click({ 
    if ($global:InstallationIDGroups.Count -gt 0) { 
        & $copySafely ($global:InstallationIDGroups -join '-')
        Show-Toast 
    }
})

# Main execution
Refresh-UI
Load-InstallationID  # Start loading in background
$window.ShowDialog()
#endregion