<#
# Copyright 2022 xiangwuxw
# MIT License. Feel free to modify and reuse. 
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

#
# How to run 
# 1. Have one iso in the folder of the script
# 2. Plug in an USB < 32G
# 3. Powershell -ExecutionPolicy Bypass -File winiso2usb.ps1
# 

#
# Performance consideration
# Use USB3 fast USB disk. Prefer > 40MBPS. 
#

#
# This PowerShell Script helps admin and technician to build bootable USB disk from ISO media
# Features include the following 
# 1. Build Legacy/UEFI dual boot USB disk, and overcome the >4G install.wim issue 
# 2. ToDo: Optional: Inject the drivers, especially the storage and network drivers, so admin/tech doesn't need to stop looking for drivers while installing. 
# 3. ToDo: Optional: Brute force remove inbox drivers so that injected OEM driver will be used in some cases.
# 4. ToDo: Optional: Inject the LCU, refresh Windows installation updated. 
# 5. ToDo: Optional: Build an updated ISO 
#

<#   ###########################################################
#    ISSSUE RESOLVED IN THIS SCRIPT: WIM FILE > 4G and UEFI Boot. 
# 
# 1. Though some UEFI implementation may support NTFS, most of the them only support FAT/FAT32, 
# 2. FAT32 max file size is 4G,
# 3. Windows WIM file is too big, even the install.wim on retail DVD may > 4G
# 4. Most captured WIM file are super large, sometime > 10G
# 
#    ######## 
#    SOLUTION
#
# 1. Create two partitions in the USB disk, first FAT32 for UEFI Boot, boot.wim to load the WINPE or Setup Environment
# 2. Windows Setup will look for the install.wim on all volumes, so put the install.wim on the second NTFS partition
# 3. Though UEFI doesn't check active partition, that set NTFS partition active will allow legacy BIOS to boot directly.  
#    In this script, set the FAT32 active, so no matter UEFI or Legacy will boot from the same code/file FAT32 path to make it easier to maintain the code.
#>

#
# Modify the policy to prevent annoying shell/explorer pop-ups. Once done, remove the setting from registry
#
# New-ItemProperty -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer -Name NoDriveTypeAutoRun  -value 255 -type Dword
# Remove-ItemProperty -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer -Name NoDriveTypeAutoRun

#
# Predefined iso file location
#
# $isofile = "c:\temp\de-de_windows_server_2022_updated_jan_2022_x64_dvd_71b18f9b.iso"

#
# Define the fat32 partition size
#
$fat32sizeG = 1
$maxusbsizeG = 32
$settingfile = "iso2usb.cfg"
$dvdtestfile = ":\efi\boot\bootx64.efi"

#
# Calculate the size in byte. 
#
$sizeinG = 1024*1024*1024
$fat32size = $fat32sizeG * $sizeinG
$maxusbsize = $maxusbsizeG * $sizeinG

#
# Check DVD ROM. If there is a Setup DVD, we will use it
# We just need one DVD
#
Get-Volume | where { ($_.DriveType -eq "CD-ROM") -and ($_.Size -gt 0) } -OutVariable dvdvol
if (!($dvdvol -is [array])) {
	if (!([string]::IsNullOrWhiteSpace($dvdvol.DriveLetter)))
	{ 
		$dvddriveletter = $dvdvol.DriveLetter.ToString()
		$dvdtestfilepath = $dvddriveletter + $dvdtestfile

		# If bootx64.efi is not there, it's not the setup disk.
		if (!(Test-path -Path $dvdtestfilepath))
		{
			# Not able to access the bootx64.efi
			$dvddriveletter = $null
		} else {
			# If use DVD, set $isofile to $null
			$isofile = $null
		}
	} 
} else {
	write-hsot "More then two DVD ROM! We can only use 1!"
}

if ([string]::IsNullOrWhiteSpace($dvddriveletter))
{
	#
	# Check the local path, 
	# 1. If there is one ISO file, use that file
	# 
	$currentpath = Get-Location 
	$currentpath = $currentpath.Path + "\*"
	Get-ChildItem -Path $currentpath -Attributes !D -Include ('*.iso', '*.img') -OutVariable isofile

	#
	# It's not just one ISO file, reset the var to null
	#
	if ($isofile.count -ne 1)
	{ 
		$isofile = $null
	} else {
		$isofile = $isofile.VersionInfo.FileName
	}	

	#
	# 2. If there is a saved setting, use the setting.
	#
	if (Test-path -Path $settingfile)
	{ 
		$tempisofile = Get-Content $settingfile 
		if (!(Test-path -Path $tempisofile))
		{
			write-host "Invalid Content in Configuration File!"
		} else {
			$isofile = $tempisofile
		}
	}	

	#
	# If the isofile is not defined, pop up a Window to allow the user to choose an ISO/IMG file. 
	#
	if ([string]::IsNullOrWhiteSpace($isofile))
	{
		Add-Type -AssemblyName System.Windows.Forms
		$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        	InitialDirectory = Get-Location
        	Filter = 'Choose a Widnows Setup ISO|*.img;*.iso'
		}
		$FileBrowser.ShowDialog()
		$isofile = $FileBrowser.FileName
	}

	#
	# If user didn't choose anything, and isofile is still null, let's break
	#
	if ([string]::IsNullOrWhiteSpace($isofile))
	{
		write-host "No ISO Found! Please rerun and choose a Windows Setup ISO file." 
		break
	}

	#
	# If we can't access the ISO file, report it
	#
	if (!(Test-path -Path $isofile))
	{ 

		write-host "ISO File is not accessible! Please rerun and choose a valid Windows Setup ISO file." 
		break

	}	

	# Save the ISO file path so that user don't have choose again next time.
	$isofile | Out-File $settingfile
}

# Get local disks and Find all USB Disks. Use the *usbstor* as a filter here unless Microsoft change the behavior
Get-Disk | where{ $_.Path -like "*usbstor*" } -OutVariable usbdisk

#
# If no USB disk found, warning and break
#
If ($usbdisk.count -eq 0) 
{
	write-host "No USB Disk Found!" 
	break	
}

#
# If 
#   1. there are more than 1 USB disk 
#   2. size > maxusbsize 32 G if we have 1, 
# Then  
#    ask user to verify the USB to be used. 
#
If( ($usbdisk.count -gt 1) -or ($usbdisk.Size -gt $maxusbsize) )
{
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing
 
	$form = New-Object System.Windows.Forms.Form
	$form.Text = 'Select a Disk'
	$form.Size = New-Object System.Drawing.Size(300,200)
	$form.StartPosition = 'CenterScreen'
 
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Location = New-Object System.Drawing.Point(75,120)
	$okButton.Size = New-Object System.Drawing.Size(75,23)
	$okButton.Text = 'OK'
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $okButton
	$form.Controls.Add($okButton)
 
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Location = New-Object System.Drawing.Point(150,120)
	$cancelButton.Size = New-Object System.Drawing.Size(75,23)
	$cancelButton.Text = 'Cancel'
	$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $cancelButton
	$form.Controls.Add($cancelButton)
 
	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Point(10,20)
	$label.Size = New-Object System.Drawing.Size(280,20)
	$label.Text = 'Please select a disk:'
	$form.Controls.Add($label)
 
	$listBox = New-Object System.Windows.Forms.ListBox
	$listBox.Location = New-Object System.Drawing.Point(10,40)
	$listBox.Size = New-Object System.Drawing.Size(260,20)
	$listBox.Height = 80
 
	$listBox.Items.Clear()
 
	#
	# Go through the disk list and add them to the list so that user can choose
	#
	foreach ($disk in $usbdisk)
	{ 
  		$diskname = $disk.Number.ToString() + ":" + [math]::Round($disk.Size/($sizeinG)).ToString() + "G:" + $disk.Model.ToString()
  		[void] $listBox.Items.Add($diskname)
	}
  
	#
	# Create the control and show it
	#
	$form.Controls.Add($listBox)
 	$form.Topmost = $true
 	$result = $form.ShowDialog()
 
	#
	# User has chosen one
	#
	if ($result -eq [System.Windows.Forms.DialogResult]::OK)
	{
    	$disknumber = $listBox.SelectedItem.Split(":")
		# Get the selected usb disk
    	$usbdisk = Get-Disk -Number $disknumber[0]
	} else {
		write-host "No USB Disk is chosen!" 
		break	
	}
}

#
# We Should have the ISO file and one USB disk here. 
#

# Clear the Disk
$usbdisk | Clear-Disk -RemoveData -confirm:$false

# Convert to MBR if it's GPT
if ($usbdisk.PartitionStyle -ne "MBR")
{
    $usbdisk | Set-Disk -PartitionStyle MBR
}

# Create first fat32 partition
$usbdisk | New-Partition -Size $fat32size

# Create t he ntfs partition
$usbdisk | New-Partition -UseMaximumSize

#
# Format the FAT32 partition and set it active and auto assign a drive letter
#
$usbdisk | Get-Partition -PartitionNumber 1 -OutVariable fat32part | Format-Volume -FileSystem FAT32 -NewFileSystemLabel "FAT32BOOT" -OutVariable fat32vol
$usbdisk | Get-Partition -PartitionNumber 1 | Set-Partition -IsActive $true
$fat32part | Add-PartitionAccessPath -AssignDriveLetter -OutVariable fat32accesspath

#
# Get the partition again to refresh the assigned letter
#
$fat32part = $fat32part | Get-Partition
$fat32driveletter = $fat32part.DriveLetter.ToString()
if ([string]::IsNullOrWhiteSpace($fat32driveletter))
{
	write-host "FAT32 Partition Initialization Failed!" 
	break	
}

#
# Format the NTFS partition and assign the drive letter
#
$usbdisk | Get-Partition -PartitionNumber 2 -OutVariable ntfspart | Format-Volume -FileSystem NTFS -NewFileSystemLabel "NTFSDATA" -OutVariable ntfs32vol
$ntfspart | Add-PartitionAccessPath -AssignDriveLetter -OutVariable ntfsaccesspath

#
# Get the partition again to refresh the assigned letter
#
$ntfspart = $ntfspart | Get-Partition
$ntfsdriveletter = $ntfspart.DriveLetter.ToString()
if ([string]::IsNullOrWhiteSpace($ntfsdriveletter))
{
	write-host "NTFS Partition Initialization Failed!" 
	break	
}

#
# Set ntfsroot and fat32root
#
$ntfsroot = $ntfsdriveletter + ":\"
$fat32root = $fat32driveletter + ":\"
$ntfsrootlog = $ntfsroot + $settingfile
$fat32rootlog = $fat32root + $settingfile

#
# Mount the ISO to virtual DVD drive and get the driver letter if we don't have a DVD
# Set the srcdriveletter here.
#
if ([string]::IsNullOrEmpty($dvddriveletter))
{
	#
	# If we don't have DVD, we use ISO
	#
	$isoimage = Mount-DiskImage -ImagePath $isofile
	$isovol = $isoimage | Get-Volume
	$isodriveletter = $isovol.DriveLetter.ToString()

	if ([string]::IsNullOrWhiteSpace($isodriveletter))
	{
		write-host "ISO DVD Mount Failed!" 
		break	
	}

	$srcdriveletter= $isodriveletter 

	#
	# Save the info to the USB so that later user can tell where it comes from.
	#
	$isofile | Out-File $ntfsrootlog
	$isofile | Out-File $fat32rootlog 

} else { 
	# Else we use DVD ROM
	$srcdriveletter = $dvddriveletter
	
	#
	# Save the info to the USB so that later user can tell where it comes from.
	#
	$dvdvol.FileSystemLabel | Out-File $ntfsrootlog
	$dvdvol.FileSystemLabel | Out-File $fat32rootlog 
}

#
# Prepare the src vars for robocopy to copy files
#
$srcroot = $srcdriveletter + ":\"
$srcbootwimfolder = $srcdriveletter + ":\sources"
$fat32bootwimfolder = $fat32driveletter + ":\sources"

#
# Use Rococopy to copy the Windows Setup Files to NTFS and FAT32 partitions
# Don't expect the copy will fail
#
robocopy $srcroot $ntfsroot /E
robocopy $srcroot $fat32root /E /XD sources DS support upgrade
robocopy $srcbootwimfolder $fat32bootwimfolder boot.wim

#
# Dismount the ISO image, eject the virtual DVD. Keep retrying every 1 second
# Not applicable for real DVD ROM
#
if ([string]::IsNullOrWhiteSpace($dvddriveletter))
{ 
	$loopcount = 0
	while($isoimage.Attached) 
	{
		$isoimage  = Dismount-DiskImage $isofile 
		if($isoimage.Attached) 
		{ 
			Start-Sleep 1
			Write-Host "Ejecting DVD Retrying Every 1 Second!"
		}

		$loopcount++
		if ($loopcount -gt 30) 
		{
			Write-Host "Failed to Eject ISO/DVD!"
			break
		}
	} 
}

# Done
write-host "Done!" 

#
# End of the file
#
