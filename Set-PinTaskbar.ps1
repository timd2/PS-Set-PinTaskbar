Function Set-PinTaskbar {
    Param (
        [Parameter(Mandatory=$True,
        HelpMessage="Target item to pin",
        ParameterSetName="Pin")]
        [string] $Pin
        ,
        [Parameter(Mandatory=$True,
        HelpMessage="Target item to unpin",
        ParameterSetName="Unpin")]
        [string] $Unpin
    )
    
    # Determine if the path specified to the pin or unpin application is valid
    If ([string]::IsNullOrEmpty($Unpin)) {
        If (!(Test-Path $Pin)) {
            Write-Warning "$Pin does not exist"
            Break
        }
        $Target = $Pin
        $PinFlag = $True
        # Get the drive letter for the application that is to be pinned or unpinned
        $DriveLetter = ($Pin -split ":")[0]
    }
    Else {
        $Target = $Unpin
        If (!(Test-Path $Unpin)) {
            Write-Warning "$Unpin does not exist"
            Break
        }
        # Get the drive letter for the application that is to be pinned or unpinned
        $DriveLetter = ($Unpin -split ":")[0]
    }

    $Reg = @{}
    $Reg.Key1 = "*"
    $Reg.Key2 = "shell"
    $Reg.Key3 = "{:}"
    $Reg.Value = "ExplorerCommandHandler"
    $Reg.Data = (Get-ItemProperty ("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\Windows.taskbarpin")).ExplorerCommandHandler
    $Reg.Path1 = "HKCU:\SOFTWARE\Classes"
    $Reg.Path2 = Join-Path $Reg.Path1 $Reg.Key1
    $Reg.Path3 = Join-Path $Reg.Path2 $Reg.Key2
    $Reg.Path4 = Join-Path $Reg.Path3 $Reg.Key3

    If (!(Test-Path -LiteralPath $Reg.Path2)) {
        # New-Item -ItemType Directory -Path $Reg.Path1 -Name [System.Management.Automation.WildcardPattern]::Escape("*") 2>$null | Out-Null
        New-Item -ItemType Directory -Path $Reg.Path1 -Name $Reg.Key1 2>$null | Out-Null
    }
    If (!(Test-Path -LiteralPath $Reg.Path3)) {
        New-Item -ItemType Directory -Path ([System.Management.Automation.WildcardPattern]::Escape($Reg.Path2)) -Name $Reg.Key2 2>&1 | Out-Null
    }
    If (!(Test-Path -LiteralPath $Reg.Path4)) {
        New-Item -ItemType Directory -Path ([System.Management.Automation.WildcardPattern]::Escape($Reg.Path3)) -Name $Reg.Key3  2>&1 | Out-Null
    }
    Set-ItemProperty -Path ([System.Management.Automation.WildcardPattern]::Escape($Reg.Path4)) -Name $Reg.Value -Value $Reg.Data  2>&1 | Out-Null

    $Shell = New-Object -ComObject "Shell.Application"
    $Folder = $Shell.Namespace((Get-Item $Target).DirectoryName)
    $Item = $Folder.ParseName((Get-Item $Target).Name)

    # Registry key where the pinned items are located
    $RegistryKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
    # Binary registry value where the pinned items are located
    $RegistryValue = "FavoritesResolve"
    # Gets the contents into an ASCII format
    $CurrentPinsProperty = ([system.text.encoding]::ASCII.GetString((Get-ItemProperty -Path $RegistryKey -Name $RegistryValue | Select-Object -ExpandProperty $RegistryValue)))
    # Filters the results for only the characters that we are looking for, so that the search will function
    [string]$CurrentPinsResults = $CurrentPinsProperty -Replace '[^\x20-\x2f^\x30-\x3a\x41-\x5c\x61-\x7F]+', ''

    # Globally Unique Identifiers for common system folders, to replace in the pin results
    $Guid = @{}
    $Guid.FOLDERID_ProgramFilesX86 = @{
        "ID" = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"
        "Path" = ${env:ProgramFiles(x86)}
    }
    $Guid.FOLDERID_ProgramFilesX64 = @{
        "ID" = "{6D809377-6AF0-444b-8957-A3773F02200E}"
        "Path" = $env:ProgramFiles
    }
    $Guid.FOLDERID_ProgramFiles = @{
        "ID" = "{905e63b6-c1bf-494e-b29c-65b732d3d21a}"
        "Path" = $env:ProgramFiles
    }
    $Guid.FOLDERID_System = @{
        "ID" = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
        "Path" = Join-Path $env:WINDIR "System32"
    }
    $Guid.FOLDERID_Windows = @{
        "ID" = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"
        "Path" = $env:WINDIR
    }

    # Replace GUIDs with full paths to the folders
    ForEach ($GuidEntry in $Guid.Keys) {
        $CurrentPinsResults = $CurrentPinsResults -replace $Guid.$GuidEntry.ID,$Guid.$GuidEntry.Path
    }

    $Split = $CurrentPinsResults -split ($env:SystemDrive)

    $SplitOutput = @()
    # Process each path entry, remove invalid characters, test to determine if the path is valid
    ForEach ($Entry in $Split) {
        If ($Entry.Substring(0,1) -eq '\') {
            # Get a list of invalid path characters
            $InvalidPathCharsRegEx = [IO.Path]::GetInvalidPathChars() -join ''
            $InvalidPathChars = "[{0}]" -f [RegEx]::Escape($InvalidPathCharsRegEx)
            $EntryProcessedPhase1 = "C:" + ($Entry -replace $InvalidPathChars)
            $EntryProcessedPhase2 = $null
            # Remove characters from the path until it is resolvable
            ForEach ($Position in $EntryProcessedPhase1.Length .. 1) {
                If (Test-Path $EntryProcessedPhase1.Substring(0,$Position)) {
                    $EntryProcessedPhase2 = $EntryProcessedPhase1.Substring(0,$Position)
                    Break
                }
            }
            # If the path resolves, add it to the array of paths
            If ($EntryProcessedPhase2) {
                $SplitOutput += $EntryProcessedPhase2
            }
        }
    }

    $PinnedItems = @()
    $Shell = New-Object -ComObject WScript.Shell
    ForEach ($Path in $SplitOutput) {
        # Determines if the entry in the registry is a link in the standard folder, if it is, resolve the path of the shortcut and add it to the array of pinnned items
        If ((Split-Path $Path) -eq (Join-Path $env:USERPROFILE "AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")) {
            $PinnedItems += $Shell.CreateShortcut($Path).TargetPath
        }
        Else {
            # If the link or executable is not in the taskbar folder, add it directly
            $PinnedItems += $Path
        }
    }
    
    # Unpin if the application is pinned
    If (!($PinFlag)) {
        If ($PinnedItems -contains $Target) {
            $Item.InvokeVerb("{:}")
            Write-Host "Unpinning application $Target"
        }
    }
    Else {
        # Only pin the application if it hasn't been pinned
        If ($PinnedItems -notcontains $Target) {
            $Item.InvokeVerb("{:}")
            Write-Host "Pinning application $Target"
        }
    }
    
    # Remove the registry key and subkeys required to pin the application
    If (Test-Path $Reg.Path3) {
        Remove-Item -LiteralPath $Reg.Path3 -Recurse 2>&1 | Out-Null
    }
}
