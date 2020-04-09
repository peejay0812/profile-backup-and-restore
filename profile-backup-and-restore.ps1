#region IMPORTS
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
Add-Type -AssemblyName System.Windows.Forms
#endregion

function Open-MainForm {
    $syncHash = [hashtable]::Synchronized(@{})
    $newRunspace = [runspacefactory]::CreateRunspace()
    $syncHash.Runspace = $newRunspace
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"
    $newRunspace.Name = "mainForm"
    $data = $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $psCmd = [powershell]::Create().AddScript({
        [xml]$xaml = @"
        <Window
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="Backup/Restore Tool" Height="185" Width="485" MaxHeight="185" MaxWidth="485" MinHeight="185" MinWidth="485" WindowStartupLocation="CenterScreen">
            <Grid>
                <Label Name="lblStatus" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" />
                <Button Name="btnBackup" Content="Backup" HorizontalAlignment="Left" Margin="231,106,0,0" VerticalAlignment="Top" Width="100" />
                <Button Name="btnRestore" Content="Restore" HorizontalAlignment="Left" Margin="349,106,0,0" VerticalAlignment="Top" Width="100"/>
                <ProgressBar Name="prgStat" HorizontalAlignment="Left" Height="23" Margin="16,35,0,0" VerticalAlignment="Top" Width="433" Minimum="0" Maximum="100"/>
                <ProgressBar Name="prgMain" HorizontalAlignment="Left" Height="23" Margin="16,70,0,0" VerticalAlignment="Top" Width="433" Minimum="0" Maximum="3" FontSize="10"/>
                <RadioButton Name="radDesk" Content="Desktop" FontSize="9px" VerticalContentAlignment="Center" HorizontalAlignment="Left" Margin="18,109,0,0" VerticalAlignment="Top"/>
                <RadioButton Name="radLap" Content="Laptop" FontSize="9px" VerticalContentAlignment="Center" HorizontalAlignment="Left" Margin="89,109,0,0" VerticalAlignment="Top"/>
            </Grid>
        </Window>
"@
        #region INITIALIZATION
        #---XAML parser---#
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $syncHash.Window = [Windows.Markup.XamlReader]::Load($reader)
        $xaml.SelectNodes("//*[@Name]") | % {$syncHash."$($_.Name)" = $syncHash.Window.FindName($_.Name)}
        #endregionW

        #region FUNCTIONS
        function Reset-Controls {
            $syncHash.MainProgress = 0
            $syncHash.Enable = $true
            $syncHash.Status = "Ready"
            $syncHash.Progress = 0
            $syncHash.HasExited = 2
            $syncHash.Reset = $false
            $syncHash.Max = 100
            $syncHash.radDesk.IsChecked = $true
        }
        #endregion

        #region CONTROLS UPDATER
        Reset-Controls
        $updateControls = {
            $syncHash.lblStatus.Content = $syncHash.Status
            $syncHash.prgStat.Value = $syncHash.Progress
            $syncHash.prgMain.Value = $syncHash.MainProgress
            $syncHash.prgStat.Maximum = $syncHash.Max

            $syncHash.btnRestore.IsEnabled = $syncHash.Enable
            $syncHash.btnBackup.IsEnabled = $syncHash.Enable
            $syncHash.radDesk.IsEnabled = $syncHash.Enable
            $syncHash.radLap.IsEnabled = $syncHash.Enable

            if($syncHash.Reset) {
                Reset-Controls
            }
        }

        $syncHash.Window.Add_SourceInitialized({
            $timer = New-Object System.Windows.Threading.DispatcherTimer   
            $timer.Interval = [TimeSpan]"0:0:0.01"          
            $timer.Add_Tick($updateControls)            
            $timer.Start()                       
        })
        #endregion

        #region CONTROL EVENTS
        $syncHash.Window.Add_Closing({
            $rs = Get-Runspace -Name "btnRun"
            $iMain = Get-Runspace -Name "mainForm"

            if($syncHash.stopwatch.IsRunning -eq $true) {
                $closeWindow = [System.Windows.Forms.MessageBox]::Show("The $($syncHash.Operation) is currently running. Closing the window will stop the operation. Do you want to proceed?",'Backup/Restore Tool','YesNo','Question')
                    if($closeWindow -eq "Yes") {
                        while($syncHash.Stopwatch2.IsRunning) {
                            Start-Sleep -Milliseconds 500
                        }

                        $rs.Close()
                        $rs.Dispose()
                        Get-Process -Name robocopy | Stop-Process -Force
                        [System.Windows.Forms.MessageBox]::Show("The $($syncHash.Operation) operation has been cancelled. The $($syncHash.Operation) might be incomplete.",'Backup/Restore Tool','OK','Warning')
                        $iMain.CloseAsync()
                    } else {
                        $_.Cancel = $true
                    }
            } else {
                $rs.Close()
                $rs.Dispose()
                $iMain.CloseAsync()
            }
        })

        $syncHash.btnBackup.Add_Click({
            $syncHash.Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $syncHash.Operation = "backup"
            $syncHash.Enable = $false
            if($syncHash.radDesk.IsChecked -eq $true) {
                $syncHash.destDir = "_BACKUP"
            } else {
                $syncHash.destDir = "_LaptopBackup"
            }

            $btnRunspace = [runspacefactory]::CreateRunspace()
            $btnRunspace.ApartmentState = "STA"
            $btnRunspace.ThreadOptions = "ReuseThread"
            $btnRunspace.Name = "btnRun"
            $btnRunspace.Open()
            $btnRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
            $cmdBtn = [powershell]::Create().AddScript({
                $syncHash.Enable = $false

                #region HELPER FUNCTIONS
                function Update-Status {
                    param (
                        [int]$Copied,
                        [int]$Count,
                        [double]$BCopied,
                        [double]$BTotal,
                        [double]$Progress,
                        [string]$Text,
                        [string]$Folder,
                        [switch]$Label
                    )

                    $convCopied = [math]::Round($BCopied/1GB, 2)
                    $convTotal = [math]::Round($BTotal/1GB, 2)
                    if($PSBoundParameters["Label"]) {
                        $syncHash.Status = $Text
                    } else {
                        $syncHash.Status = "Backing up $Folder - copied $Copied of $Count files - $($convCopied)GB/$($convTotal)GB"
                        $syncHash.Progress = $Progress
                    }
                }

                function Copy-Files {
                    [CmdletBinding()]
                    param (
                        [string]$Source,
                        [string]$Destination,
                        [string]$Folder,
                        [int]$FolderCount,
                        [int]$ReportGap = 1000
                    )
                    
                    $RegexBytes = '(?<=\s+)\d+(?=\s+)';
                    $CommonRobocopyParams = '*.* /E /NP /NDL /NC /BYTES /NJH /NJS /R:3 /W:3 /MT:32';
                
                    #region Robocopy Staging
                    Update-Status -Text "Preparing backup..." -Label
                    $syncHash.Stopwatch2 = [System.Diagnostics.Stopwatch]::StartNew()
                    $StagingLogPath = '{0}\temp\{1} robocopy staging.log' -f $env:windir, (Get-Date -Format 'yyyy-MM-dd HH-mm-ss');
                    $StagingArgumentList = '"{0}" "{1}" /LOG:"{2}" /L {3}' -f $Source, $Destination, $StagingLogPath, $CommonRobocopyParams;
                    Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList -WindowStyle Hidden;
                    $StagingContent = Get-Content -Path $StagingLogPath;
                    $TotalFileCount = $StagingContent.Count - 1;
                    [RegEx]::Matches(($StagingContent -join "`n"), $RegexBytes) | % { $BytesTotal = 0; } { $BytesTotal += $_.Value; };
                    #endregion Robocopy Staging
                
                    #region Start Robocopy
                    $RobocopyLogPath = '{0}\temp\{1} robocopy.log' -f $env:windir, (Get-Date -Format 'yyyy-MM-dd HH-mm-ss');
                    $ArgumentList = '"{0}" "{1}" /LOG:"{2}" {3}' -f $Source, $Destination, $RobocopyLogPath, $CommonRobocopyParams;
                    $syncHash.MainProgress = $FolderCount
                    $Robocopy = Start-Process -FilePath robocopy.exe -ArgumentList $ArgumentList -Verbose -PassThru -WindowStyle Hidden;
                    $syncHash.Stopwatch2.Stop()
                    Start-Sleep -Milliseconds 100;
                    #endregion Start Robocopy
                
                    #region Progress bar loop
                    while (!$Robocopy.HasExited) {
                        Start-Sleep -Milliseconds $ReportGap;
                        $BytesCopied = 0;
                        $LogContent = Get-Content -Path $RobocopyLogPath;
                        $BytesCopied = [Regex]::Matches($LogContent, $RegexBytes) | ForEach-Object -Process { $BytesCopied += $_.Value; } -End { $BytesCopied; };
                        $CopiedFileCount = $LogContent.Count - 1;
                        
                        $Percentage = 0;
                        if ($BytesCopied -gt 0) {
                           $Percentage = (($BytesCopied/$BytesTotal)*100)
                        }

                        Update-Status -Copied $CopiedFileCount -Count $TotalFileCount -BCopied $BytesCopied -BTotal $BytesTotal -Progress $Percentage -Folder $Folder
                    }
                    #endregion Progress loop
                }
                #endregion
                
                $dirList = @("Printers", "Desktop", "Documents", "Downloads", "Favorites", "Music", "Pictures", "Videos", "AppData\Local\Google\Chrome\User Data\Default", "AppData\Local\Mozilla\Firefox", "AppData\Roaming\Mozilla\Firefox\Profiles")
                $backupDrive = "H:"
                $folderCount = 1

                $accept = [System.Windows.Forms.MessageBox]::Show("Proceeding will backup your files to the $backupDrive drive under a folder named $($syncHash.destDir).`n`nIf you do not want to proceed press Cancel. Press OK to proceed with backup.",'Backup/Restore Tool','OKCancel','Information')                
                if($accept -eq "OK") {
                    if((Test-Path -Path $backupDrive)) {
                        #---Backup printers---#
                        $printerPath = "$env:USERPROFILE\Printers\printerlist.txt"
                        if(Test-Path -Path $printerPath) {
                            Remove-Item -Path $printerPath
                        }

                        New-Item -Path $printerPath -ItemType File -Force | Out-Null

                        Update-Status -Text "Retrieving printers..." -Label
                        [array]$printers = Get-WmiObject Win32_Printer | Where-Object network -eq $true | Select-Object Name, PortName
                        
                        if($printers.Count -gt 0) {
                            foreach($print in $printers) {
                                $print.Name | Out-File -FilePath $printerPath -Append
                            }
                        }

                        #---Actual backup job---#
                        foreach($dir in $dirList) {
                            $source = "$env:USERPROFILE\$dir"
                            $destination = "$backupDrive\$($syncHash.destDir)\$env:USERNAME\$dir"
                            
                            if((Test-Path -Path $source)) {
                                Copy-Files -Source $source -Destination $destination -Folder $dir -FolderCount $folderCount
                            }
                            
                            $folderCount++
                        }
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("There was an error connecting to backup destination. Please make sure you are connected to the network.",'Backup/Restore Tool','OK','Warning')
                        $syncHash.Enable = $true
                        $syncHash.HasExited = 0
                        break
                    }
                    
                    if($syncHash.HasExited -eq 2) {
                        [System.Windows.Forms.MessageBox]::Show("Backup complete. Please check to confirm files were backed up",'Backup/Restore Tool','OK','Information')
                        Invoke-Item "$backupDrive\$($syncHash.destDir)\$env:USERNAME"
                    }
                }
                $syncHash.Stopwatch.Stop()
                $syncHash.Reset = $true
            })
            $cmdBtn.Runspace = $btnRunspace
            $cmdBtn.BeginInvoke() | Out-Null
        })

        $syncHash.btnRestore.Add_Click({
            $syncHash.Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $syncHash.Operation = "restore"
            $syncHash.Enable = $false
            if($syncHash.radDesk.IsChecked -eq $true) {
                $syncHash.rad = "desktop"
                $syncHash.srcDir = "_BACKUP"
            } else {
                $syncHash.rad = "laptop"
                $syncHash.srcDir = "_LaptopBackup"
            }

            $btnRunspace = [runspacefactory]::CreateRunspace()
            $btnRunspace.ApartmentState = "STA"
            $btnRunspace.ThreadOptions = "ReuseThread"
            $btnRunspace.Name = "btnRun"
            $btnRunspace.Open()
            $btnRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
            $cmdBtn = [powershell]::Create().AddScript({
                $syncHash.Enable = $false

                #region HELPER FUNCTIONS
                function Update-Status {
                    param (
                        [int]$Copied,
                        [int]$Count,
                        [double]$BCopied,
                        [double]$BTotal,
                        [double]$Progress,
                        [string]$Text,
                        [string]$Folder,
                        [switch]$Label,
                        [switch]$Printer
                    )

                    $convCopied = [math]::Round($BCopied/1GB, 2)
                    $convTotal = [math]::Round($BTotal/1GB, 2)
                    if($PSBoundParameters["Label"]) {
                        $syncHash.Status = $Text
                    } elseif($PSBoundParameters["Printer"]) {
                        $syncHash.Status = $Text
                        $syncHash.Progress = $Progress
                    } else {
                        $syncHash.Status = "Restoring $Folder - copied $Copied of $Count files - $($convCopied)GB/$($convTotal)GB"
                        $syncHash.Progress = $Progress
                    }
                    
                }

                function Copy-Files {
                    [CmdletBinding()]
                    param (
                        [string]$Source,
                        [string]$Destination,
                        [string]$Folder,
                        [int]$FolderCount,
                        [int]$ReportGap = 1000
                    )
                    
                    $RegexBytes = '(?<=\s+)\d+(?=\s+)';
                    $CommonRobocopyParams = '*.* /E /NP /NDL /NC /BYTES /NJH /NJS /R:3 /W:3 /MT:32';
                
                    #region Robocopy Staging
                    Update-Status -Text "Preparing restore..." -Label
                    $syncHash.Stopwatch2 = [System.Diagnostics.Stopwatch]::StartNew()
                    $StagingLogPath = '{0}\temp\{1} robocopy staging.log' -f $env:windir, (Get-Date -Format 'yyyy-MM-dd HH-mm-ss');
                    $StagingArgumentList = '"{0}" "{1}" /LOG:"{2}" /L {3}' -f $Source, $Destination, $StagingLogPath, $CommonRobocopyParams;
                    Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList -WindowStyle Hidden;
                    $StagingContent = Get-Content -Path $StagingLogPath;
                    $TotalFileCount = $StagingContent.Count - 1;
                    [RegEx]::Matches(($StagingContent -join "`n"), $RegexBytes) | % { $BytesTotal = 0; } { $BytesTotal += $_.Value; };
                    #endregion Robocopy Staging
                
                    #region Start Robocopy
                    $RobocopyLogPath = '{0}\temp\{1} robocopy.log' -f $env:windir, (Get-Date -Format 'yyyy-MM-dd HH-mm-ss');
                    $ArgumentList = '"{0}" "{1}" /LOG:"{2}" {3}' -f $Source, $Destination, $RobocopyLogPath, $CommonRobocopyParams;
                    $syncHash.MainProgress = $FolderCount
                    $Robocopy = Start-Process -FilePath robocopy.exe -ArgumentList $ArgumentList -Verbose -PassThru -WindowStyle Hidden;
                    $syncHash.Stopwatch2.Stop()
                    Start-Sleep -Milliseconds 100;
                    #endregion Start Robocopy
                
                    #region Progress bar loop
                    while (!$Robocopy.HasExited) {
                        Start-Sleep -Milliseconds $ReportGap;
                        $BytesCopied = 0;
                        $LogContent = Get-Content -Path $RobocopyLogPath;
                        $BytesCopied = [Regex]::Matches($LogContent, $RegexBytes) | ForEach-Object -Process { $BytesCopied += $_.Value; } -End { $BytesCopied; };
                        $CopiedFileCount = $LogContent.Count - 1;
                        
                        $Percentage = 0;
                        if ($BytesCopied -gt 0) {
                           $Percentage = (($BytesCopied/$BytesTotal)*100)
                        }

                        Update-Status -Copied $CopiedFileCount -Count $TotalFileCount -BCopied $BytesCopied -BTotal $BytesTotal -Progress $Percentage -Folder $Folder
                    }
                    #endregion Progress loop
                }
                #endregion
                
                $dirList = @("Printers", "Desktop", "Documents", "Downloads", "Favorites", "Music", "Pictures", "Videos", "AppData\Local\Google\Chrome\User Data\Default", "AppData\Local\Mozilla\Firefox", "AppData\Roaming\Mozilla\Firefox\Profiles")
                $backupDrive = "H:"
                $folderCount = 1

                $accept = [System.Windows.Forms.MessageBox]::Show("A restore should only be done in the case that you received a new computer, or your computer has crashed. This process will restore the last backup you made from your previous $($syncHash.rad). If you do not want to proceed press Cancel. Press OK to proceed with restore.",'Backup/Restore Tool','OKCancel','Information')
                if($accept -eq "OK") {
                    if((Test-Path -Path $backupDrive)) {
                        if((Test-Path -Path "$backupDrive\$($syncHash.srcDir)\$env:USERNAME")) {
                            foreach($dir in $dirList) {
                                $destination = "$env:USERPROFILE\$dir"
                                $source = "$backupDrive\$($syncHash.srcDir)\$env:USERNAME\$dir"
                                
                                if((Test-Path -Path $source)) {
                                    Copy-Files -Source $source -Destination $destination -Folder $dir -FolderCount $folderCount
                                    $backupValid = $true
                                }
                                
                                $folderCount++
                            }
            
                            if($syncHash.HasExited -eq 2 -and $backupValid -eq $true) {
                                $printerPath = "$env:USERPROFILE\Printers\printerlist.txt"
                                $printers = Get-Content -Path $printerPath
                                if($printers.Count -gt 0) {
                                    $syncHash.Max = $printers.Count
                                    $printCount = 1
                                    foreach($print in $printers) {
                                        Update-Status -Text "Restoring printer - $print..." -Progress $printCount -Printer
                                        Invoke-Command -ScriptBlock {RUNDLL32 PRINTUI.DLL,PrintUIEntry /in /n $print} -ErrorAction SilentlyContinue
                                        $printCount++
                                    }
                                }
    
                                [System.Windows.Forms.MessageBox]::Show("Restore complete!",'Backup/Restore Tool','OK','Information')
                            }
                        } else {
                            [System.Windows.Forms.MessageBox]::Show("Restore failed. Make sure you are selecting the correct backup type",'Backup/Restore Tool','OK','Error')
                            $syncHash.Stopwatch.Stop()
                            $syncHash.Enable = $true
                            $syncHash.HasExited = 0
                            break
                        }
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("There was an error connecting to the backup source. Please make sure you are connected to the network.",'Backup/Restore Tool','OK','Warning')
                        $syncHash.Stopwatch.Stop()
                        $syncHash.Enable = $true
                        $syncHash.HasExited = 0
                        break
                    }
                }
                
                $syncHash.Stopwatch.Stop()
                $syncHash.Reset = $true
            })
            $cmdBtn.Runspace = $btnRunspace
            $cmdBtn.BeginInvoke() | Out-Null
        })
        #endregion
        
        $syncHash.Window.ShowDialog() | Out-Null
        $syncHash.Error += $Error
    })

    $psCmd.Runspace = $newRunspace
    $data = $psCmd.BeginInvoke()
    
    Return $syncHash
}

$prg = Open-MainForm

#---Post closing cleanup---#
do {
    $main = Get-Runspace -Name "mainForm"
    Start-Sleep -Milliseconds 100
} while($main.RunspaceAvailability -ne "None")

$main.Dispose()