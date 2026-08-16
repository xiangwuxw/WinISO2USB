<#
# Copyright 2026 xiangwuxw
# MIT License
# Updated: Converted to GPT, fixed boot.sdi missing issue, improved USB detection,
#          added AutoPlay safe rollback, strict FriendlyName/Size (e.g. Samsung 60G) manual confirmation,
#          and full PowerShell Write-Progress tracking (Completed/Total files).
#>

#Requires -RunAsAdministrator

#
# Performance & Settings
#
$fat32sizeG  = 1
$maxusbsizeG = 32
$settingfile  = "iso2usb.cfg"
$dvdtestfile  = ":\efi\boot\bootx64.efi"

$sizeinG    = 1GB
$fat32size  = $fat32sizeG * $sizeinG
$maxusbsize = $maxusbsizeG * $sizeinG

#
# AutoPlay Registry Path & State Tracking
#
$regPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
$regName = "NoDriveTypeAutoRun"
$script:autoPlayValueExisted = $false
$script:originalAutoPlayValue = $null
$script:isoMounted = $false
$script:isofile = $null

# ---------------------------------------------------------
# Helper Function: Copy Files with Write-Progress Bar
# ---------------------------------------------------------
function Copy-DirectoryWithProgress {
    param (
        [Parameter(Mandatory=$true)][string]$SourcePath,
        [Parameter(Mandatory=$true)][string]$DestinationPath,
        [string[]]$ExcludeDirs = @(),
        [string]$Activity = "Copying Files"
    )

    $SourcePath = $SourcePath.TrimEnd('\')
    $DestinationPath = $DestinationPath.TrimEnd('\')

    Write-Host "Calculating file list for: $Activity..." -ForegroundColor DarkGray
    
    # Enumerate and filter files
    $allFiles = Get-ChildItem -Path $SourcePath -Recurse -File | Where-Object {
        $relativePath = $_.FullName.Substring($SourcePath.Length).TrimStart('\')
        $firstDir = ($relativePath -split '\\')[0]
        
        $exclude = $false
        foreach ($ex in $ExcludeDirs) {
            if ($relativePath -like "$ex*" -or $firstDir -ieq $ex) {
                $exclude = $true
                break
            }
        }
        -not $exclude
    }

    $total = $allFiles.Count
    if ($total -eq 0) { return }

    $current = 0
    foreach ($file in $allFiles) {
        $current++
        $relativePath = $file.FullName.Substring($SourcePath.Length)
        $destFile = Join-Path $DestinationPath $relativePath
        $destDir = [System.IO.Path]::GetDirectoryName($destFile)

        if (!(Test-Path -Path $destDir)) {
            New-Item -Path $destDir -ItemType Directory -Force | Out-Null
        }

        $percent = [math]::Round(($current / $total) * 100)
        $statusText = "[{0}/{1}] ({2}%) {3}" -f $current, $total, $percent, $file.Name
        Write-Progress -Activity $Activity -Status $statusText -PercentComplete $percent

        # Native .NET copy for maximum throughput and reliability
        [System.IO.File]::Copy($file.FullName, $destFile, $true)
    }

    Write-Progress -Activity $Activity -Completed
}

# ---------------------------------------------------------
# Step 0: Backup and Disable AutoPlay
# ---------------------------------------------------------
try {
    if (Test-Path -Path $regPath) {
        $existingProp = Get-ItemProperty -Path $regPath -Name $regName -ErrorAction SilentlyContinue
        if ($null -ne $existingProp -and $null -ne $existingProp.$regName) {
            $script:autoPlayValueExisted = $true
            $script:originalAutoPlayValue = $existingProp.$regName
        }
    } else {
        New-Item -Path $regPath -Force | Out-Null
    }

    Set-ItemProperty -Path $regPath -Name $regName -Value 255 -Type DWord -Force
    Write-Host "AutoPlay temporarily disabled for disk operations." -ForegroundColor DarkGray
} catch {
    Write-Warning "Failed to modify AutoPlay registry settings: $($_.Exception.Message)"
}

# ---------------------------------------------------------
# Main Execution Block
# ---------------------------------------------------------
try {
    #
    # 1. Check for physical Setup DVD ROM
    #
    Get-Volume | Where-Object { ($_.DriveType -eq "CD-ROM") -and ($_.Size -gt 0) } -OutVariable dvdvol
    if (!($dvdvol -is [array])) {
        if (!([string]::IsNullOrWhiteSpace($dvdvol.DriveLetter))) { 
            $dvddriveletter = $dvdvol.DriveLetter.ToString()
            $dvdtestfilepath = $dvddriveletter + $dvdtestfile

            if (!(Test-Path -Path $dvdtestfilepath)) {
                $dvddriveletter = $null
            } else {
                $script:isofile = $null
            }
        } 
    } else {
        Write-Warning "More than one DVD ROM detected! Using the first available."
    }

    #
    # 2. Select ISO file if no valid physical DVD found
    #
    if ([string]::IsNullOrWhiteSpace($dvddriveletter)) {
        $currentpath = (Get-Location).Path + "\*"
        Get-ChildItem -Path $currentpath -Attributes !D -Include ('*.iso', '*.img') -OutVariable isofileResult

        if ($isofileResult.Count -eq 1) { 
            $script:isofile = $isofileResult.FullName
        }

        if (Test-Path -Path $settingfile) { 
            $tempisofile = (Get-Content $settingfile -Raw).Trim()
            if (Test-Path -Path $tempisofile) {
                $script:isofile = $tempisofile
            }
        }	

        if ([string]::IsNullOrWhiteSpace($script:isofile)) {
            Add-Type -AssemblyName System.Windows.Forms
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
                InitialDirectory = (Get-Location).Path
                Filter           = 'Windows Setup Image (*.iso;*.img)|*.iso;*.img'
                Title            = 'Select a Windows Setup ISO/IMG File'
            }
            if ($FileBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $script:isofile = $FileBrowser.FileName
            }
        }

        if ([string]::IsNullOrWhiteSpace($script:isofile) -or !(Test-Path -Path $script:isofile)) {
            throw "No valid ISO/IMG file selected or accessible!"
        }

        $script:isofile | Out-File $settingfile -Encoding utf8
    }

    #
    # 3. Detect USB Disks
    #
    $usbdisk = Get-Disk | Where-Object { $_.BusType -eq 'USB' }

    if (!$usbdisk -or $usbdisk.Count -eq 0) {
        throw "No USB Disk detected!"
    }

    $requiresStrictConfirmation = $false

    #
    # 4. User selection if multiple USBs or USB > 32GB
    #
    if (($usbdisk.Count -gt 1) -or ($usbdisk.Size -gt $maxusbsize)) {
        $requiresStrictConfirmation = $true
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
     
        $form = New-Object System.Windows.Forms.Form -Property @{
            Text            = 'Select Target USB Disk'
            Size            = New-Object System.Drawing.Size(380, 260)
            StartPosition   = 'CenterScreen'
            FormBorderStyle = 'FixedDialog'
            MaximizeBox     = $false
        }
     
        $label = New-Object System.Windows.Forms.Label -Property @{
            Location  = New-Object System.Drawing.Point(15, 12)
            Size      = New-Object System.Drawing.Size(340, 20)
            Text      = 'WARNING: Selected disk will be completely WIPED!'
            ForeColor = [System.Drawing.Color]::Red
        }
        $form.Controls.Add($label)

        $listBox = New-Object System.Windows.Forms.ListBox -Property @{
            Location = New-Object System.Drawing.Point(15, 38)
            Size     = New-Object System.Drawing.Size(335, 110)
        }
     
        foreach ($disk in $usbdisk) { 
            $firstWord = ($disk.FriendlyName.Trim() -split '\s+')[0]
            $roundedUpG = [math]::Ceiling($disk.Size / 1GB)
            $diskname = "Disk {0}: {1} {2}G ({3})" -f $disk.Number, $firstWord, $roundedUpG, $disk.FriendlyName
            [void]$listBox.Items.Add($diskname)
        }
        $listBox.SelectedIndex = 0
        $form.Controls.Add($listBox)

        $okButton = New-Object System.Windows.Forms.Button -Property @{
            Location     = New-Object System.Drawing.Point(85, 165)
            Size         = New-Object System.Drawing.Size(85, 30)
            Text         = 'OK'
            DialogResult = [System.Windows.Forms.DialogResult]::OK
        }
        $form.AcceptButton = $okButton
        $form.Controls.Add($okButton)
     
        $cancelButton = New-Object System.Windows.Forms.Button -Property @{
            Location     = New-Object System.Drawing.Point(195, 165)
            Size         = New-Object System.Drawing.Size(85, 30)
            Text         = 'Cancel'
            DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        }
        $form.CancelButton = $cancelButton
        $form.Controls.Add($cancelButton)
     
        $form.Topmost = $true
        if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedText = $listBox.SelectedItem.ToString()
            $selectedDiskNumber = [regex]::Match($selectedText, '^Disk (\d+):').Groups[1].Value
            $usbdisk = Get-Disk -Number $selectedDiskNumber
        } else {
            throw "Operation cancelled by user during disk selection."
        }
    }

    #
    # 5. Extract Details & Perform Confirmation Verification (Explicitly keeping 'G')
    #
    $targetFirstWord = ($usbdisk.FriendlyName.Trim() -split '\s+')[0]
    $targetRoundedUpG = [math]::Ceiling($usbdisk.Size / 1GB)
    $expectedConfirmation = "$targetFirstWord ${targetRoundedUpG}G"

    Write-Host "`nTarget Disk Details:" -ForegroundColor Yellow
    Write-Host "----------------------------------------" -ForegroundColor Yellow
    Write-Host "  Disk Number   : $($usbdisk.Number)"
    Write-Host "  Model / Name  : $($usbdisk.FriendlyName)"
    Write-Host "  Target Match  : $expectedConfirmation"
    Write-Host "----------------------------------------" -ForegroundColor Yellow

    if ($requiresStrictConfirmation) {
        Write-Host "Safety Check: To prevent accidental data loss, please confirm the target drive." -ForegroundColor Red
        Write-Host "Type exactly: '$expectedConfirmation' and press Enter:" -ForegroundColor Cyan
        $userInput = (Read-Host "Confirmation").Trim()

        # 兼容输入匹配：支持 "Samsung 60G"、"Samsung 60GB" 或 "Samsung 60"
        $validPatterns = @($expectedConfirmation, "$expectedConfirmation`B", "$targetFirstWord $targetRoundedUpG")
        $isMatched = $false
        foreach ($pattern in $validPatterns) {
            if ($userInput -like $pattern) {
                $isMatched = $true
                break
            }
        }

        if (!$isMatched) {
            throw "Safety confirmation mismatch! Expected '$expectedConfirmation', got '$userInput'. Formatting aborted."
        }
        Write-Host "Confirmation verified successfully.`n" -ForegroundColor Green
    }

    Write-Host ("Preparing to wipe and format Disk {0} ({1})..." -f $usbdisk.Number, $usbdisk.FriendlyName) -ForegroundColor Yellow

    #
    # 6. Wipe & Initialize as GPT
    #
    $usbdisk | Clear-Disk -RemoveData -RemoveOEM -Confirm:$false
    $usbdisk = Get-Disk -Number $usbdisk.Number

    if ($usbdisk.PartitionStyle -eq "RAW") {
        $usbdisk | Initialize-Disk -PartitionStyle GPT | Out-Null
    } elseif ($usbdisk.PartitionStyle -ne "GPT") {
        $usbdisk | Set-Disk -PartitionStyle GPT
    }

    #
    # 7. Create Partitions: FAT32 (UEFI Boot) + NTFS (Data/Install.wim)
    #
    Write-Host "Creating GPT partitions..." -ForegroundColor Cyan
    $fat32Part = $usbdisk | New-Partition -Size $fat32size -AssignDriveLetter
    $ntfsPart  = $usbdisk | New-Partition -UseMaximumSize -AssignDriveLetter

    $fat32Vol = $fat32Part | Format-Volume -FileSystem FAT32 -NewFileSystemLabel "FAT32BOOT" -Confirm:$false
    $ntfsVol  = $ntfsPart  | Format-Volume -FileSystem NTFS  -NewFileSystemLabel "NTFSDATA"  -Confirm:$false

    $fat32DriveLetter = ($fat32Part | Get-Partition).DriveLetter
    $ntfsDriveLetter  = ($ntfsPart  | Get-Partition).DriveLetter

    if ([string]::IsNullOrWhiteSpace($fat32DriveLetter) -or [string]::IsNullOrWhiteSpace($ntfsDriveLetter)) {
        throw "Failed to assign drive letters to newly created partitions."
    }

    $fat32root = "$($fat32DriveLetter):\"
    $ntfsroot  = "$($ntfsDriveLetter):\"

    #
    # 8. Mount ISO / Read Source Media
    #
    if ([string]::IsNullOrEmpty($dvddriveletter)) {
        Write-Host "Mounting ISO image..." -ForegroundColor Cyan
        $mountResult = Mount-DiskImage -ImagePath $script:isofile -PassThru
        $script:isoMounted = $true
        Start-Sleep -Seconds 2
        $isovol = $mountResult | Get-Volume
        $srcdriveletter = $isovol.DriveLetter
        
        if ([string]::IsNullOrWhiteSpace($srcdriveletter)) {
            throw "Failed to retrieve mounted ISO drive letter."
        }
    } else {
        $srcdriveletter = $dvddriveletter
    }

    $srcroot = "$($srcdriveletter):\"
    $srcsources = Join-Path $srcroot "sources"
    $fat32sources = Join-Path $fat32root "sources"

    # Record metadata
    $sourceTag = if ($dvddriveletter) { $dvdvol.FileSystemLabel } else { $script:isofile }
    $sourceTag | Out-File (Join-Path $ntfsroot $settingfile) -Encoding utf8
    $sourceTag | Out-File (Join-Path $fat32root $settingfile) -Encoding utf8

    #
    # 9. Copy Files with PowerShell Write-Progress
    #
    Write-Host "`n[Step 1/2] Copying full source files to NTFS partition..." -ForegroundColor Green
    Copy-DirectoryWithProgress -SourcePath $srcroot -DestinationPath $ntfsroot -Activity "[Step 1/2] Copying media to NTFS partition"

    Write-Host "`n[Step 2/2] Copying UEFI boot files to FAT32 partition..." -ForegroundColor Green
    Copy-DirectoryWithProgress -SourcePath $srcroot -DestinationPath $fat32root -ExcludeDirs @('sources', 'DS', 'support', 'upgrade') -Activity "[Step 2/2] Copying UEFI boot files to FAT32"

    if (!(Test-Path -Path $fat32sources)) { 
        New-Item -ItemType Directory -Path $fat32sources | Out-Null 
    }

    $bootFiles = @("boot.wim", "boot.sdi") | Where-Object { Test-Path (Join-Path $srcsources $_) }
    $totalBoot = $bootFiles.Count
    $curBoot = 0
    foreach ($bf in $bootFiles) {
        $curBoot++
        $percent = [math]::Round(($curBoot / $totalBoot) * 100)
        Write-Progress -Activity "[Step 2/2] Copying boot files (boot.wim & boot.sdi)" -Status ("[{0}/{1}] ({2}%) {3}" -f $curBoot, $totalBoot, $percent, $bf) -PercentComplete $percent
        [System.IO.File]::Copy((Join-Path $srcsources $bf), (Join-Path $fat32sources $bf), $true)
    }
    Write-Progress -Activity "[Step 2/2] Copying boot files (boot.wim & boot.sdi)" -Completed

    Write-Host "`n=======================================================" -ForegroundColor Green
    Write-Host " Bootable GPT/UEFI USB creation completed successfully! " -ForegroundColor Green
    Write-Host "=======================================================" -ForegroundColor Green

} catch {
    Write-Host "`n[ERROR] An error occurred during execution: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # ---------------------------------------------------------
    # Cleanup & Restore (Guaranteed Execution)
    # ---------------------------------------------------------
    Write-Host "`nPerforming environment cleanup..." -ForegroundColor DarkGray

    # 1. Dismount ISO if mounted
    if ($script:isoMounted -and !([string]::IsNullOrWhiteSpace($script:isofile))) {
        Write-Host "Dismounting virtual ISO image..." -ForegroundColor DarkGray
        Dismount-DiskImage -ImagePath $script:isofile -ErrorAction SilentlyContinue | Out-Null
    }

    # 2. Restore AutoPlay / NoDriveTypeAutoRun settings
    try {
        if ($script:autoPlayValueExisted) {
            Set-ItemProperty -Path $regPath -Name $regName -Value $script:originalAutoPlayValue -Type DWord -Force -ErrorAction SilentlyContinue
        } else {
            Remove-ItemProperty -Path $regPath -Name $regName -ErrorAction SilentlyContinue
        }
        Write-Host "AutoPlay registry settings reverted successfully." -ForegroundColor DarkGray
    } catch {
        Write-Warning "Failed to revert AutoPlay settings: $($_.Exception.Message)"
    }
}
