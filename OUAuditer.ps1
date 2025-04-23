# OU Auditer
# Author: Dawson Adams (dawsonaa@ksu.edu, https://github.com/dawsonaa)
# Organization: Kansas State University
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.DirectoryServices
Add-Type -AssemblyName System.Drawing

$iconPath = Join-Path $PSScriptRoot "icon.ico"
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
$currentDate = $null

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "The ImportExcel module is not installed. Running Install-Module command."
    [System.Windows.Forms.MessageBox]::Show("ImportExcel module is not installed, Attempting to install. Please type 'A' in powershell to install.", "Required Module not installed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    Install-Module -Name ImportExcel -Scope CurrentUser
}

function Get-FolderAccess {
    param (
        [string[]]$groupNames,
        [string]$folderPath,
        [int]$folderDepth = 2,
        [bool]$recursive = $false
    )

    Write-Host "`nGetting folder access for groups: $($groupNames -join ', ')`nFolder Path: $folderPath`nFolder Depth: $folderDepth"
    $startTime = [DateTime]::Now.Ticks
    $accessList = @{}
    $rootAcl = Get-Acl -Path $folderPath
    $rootGroups = $rootAcl.Access | ForEach-Object { $_.IdentityReference.Value }

    foreach ($groupName in $groupNames) {
        $accessList[$groupName] = @()
        if ($rootGroups -contains $groupName) {
            $accessList[$groupName] += [PSCustomObject]@{ Folders = $folderPath; 'Access Types' = $_.FileSystemRights }
        }
    }

    if ($folderDepth -eq 0) {
        $folders = @([PSCustomObject]@{ FullName = $folderPath })
    }
    elseif ($recursive) {
        $folders = @([PSCustomObject]@{ FullName = $folderPath }) + (Get-ChildItem -Path $folderPath -Directory -Recurse)
    }
    else {
        $folders = @([PSCustomObject]@{ FullName = $folderPath }) + (Get-ChildItem -Path $folderPath -Directory -Depth ($folderDepth - 1))
    }

    foreach ($folder in $folders) {
        if ($null -eq $folder.FullName -or $folder.FullName -eq "") {
            continue
        }

        $acl = Get-Acl -Path $folder.FullName
        $parentFolder = Get-Item -Path $folder.FullName | Select-Object -ExpandProperty Parent
        if ($null -eq $parentFolder -or $parentFolder.FullName -eq "") {
            continue
        }

        $parentAcl = Get-Acl -Path $parentFolder.FullName
        $parentGroups = $parentAcl.Access | ForEach-Object { $_.IdentityReference.Value }
        $folderGroups = $acl.Access | ForEach-Object { $_.IdentityReference.Value }
        if ($folderGroups -ne $parentGroups) {
            foreach ($groupName in $groupNames) {
                foreach ($access in $acl.Access) {
                    if ($access.IdentityReference -like "*$groupName*") {
                        $accessList[$groupName] += [PSCustomObject]@{ Folders = $folder.FullName; 'Access Types' = $access.FileSystemRights }
                        break
                    }
                }
            }
        }
    }

    $endTime = [DateTime]::Now.Ticks
    Write-Host ("Time taken to get folder access: " + (($endTime - $startTime) / 10000000) + " s`n")
    return $accessList
}

function Get-DateWithOrdinal {
    param (
        [datetime]$date
    )

    $day = $date.Day
    $suffix = switch ($day) {
        { ($_ -lt 11 -or $_ -gt 13) -and ($_ % 10 -eq 1) } { "st"; break }
        { ($_ -lt 11 -or $_ -gt 13) -and ($_ % 10 -eq 2) } { "nd"; break }
        { ($_ -lt 11 -or $_ -gt 13) -and ($_ % 10 -eq 3) } { "rd"; break }
        default { "th" }
    }

    return "{0} {1}{2}, {3} at {4}" -f `
        $date.ToString("MMMM"), `
        $day, `
        $suffix, `
        $date.Year, `
        $date.ToString("h:mm tt")
}

function Add-LegendSheet {
    param (
        [string]$excelFile,
        [string]$folderPath,
        [string]$distinguishedName
    )

    $department = $distinguishedName -split ',' | Select-Object -First 1
    $departmentName = $department -split '=' | Select-Object -Last 1
    $formattedDate = Get-DateWithOrdinal -date $currentDate

    $excelPackage = Open-ExcelPackage -Path $excelFile

    $legendSheet = $excelPackage.Workbook.Worksheets.Add("Legend")
    $legendSheet.TabColor = [System.Drawing.Color]::Green

    $legendSheet.Cells["A1"].Value = "Active Directory Organizational Unit Audit - $departmentName - $formattedDate"
    $legendSheet.Cells["A1"].Style.Font.Bold = $true

    $legendSheet.Cells["A2"].Value = "Distinguished Name"
    $legendSheet.Cells["B2"].Value = $distinguishedName
    $legendSheet.Cells["A3"].Value = "Folder Path"
    $legendSheet.Cells["B3"].Value = $folderPath
    $legendSheet.Cells["A4"].Value = "Folder Depth"
    $legendSheet.Cells["B4"].Value = $folderDepth
    $legendSheet.Cells["B4"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left

    $legendSheet.Cells["A5"].Value = ""
    $legendSheet.Cells["A6"].Value = "Instructions"
    $legendSheet.Cells["A6"].Style.Font.Bold = $true
    $legendSheet.Cells["A7"].Value = "Go through each group and use the below colors to mark groups/locations as needed."

    $legendSheet.Cells["A9"].Value = "Group Actions"
    $legendSheet.Cells["A9"].Style.Font.Bold = $true

    $legendSheet.Cells["A10"].Value = "Add user or file location to group"
    $legendSheet.Cells["A11"].Value = "Remove user or file location from group"

    $legendSheet.Cells["B10"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $legendSheet.Cells["B10"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen)

    $legendSheet.Cells["B11"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $legendSheet.Cells["B11"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Red)

    $legendSheet.Cells["A12"].Value = ""
    $legendSheet.Cells["A13"].Value = "Please provide the full file path, e.g."
    $legendSheet.Cells["A14"].Value = "\\catfiles.users.campus\workarea$\Dept\Folder\Location"

    $legendSheet.Cells["A13"].Style.Font.Bold = $true

    $legendSheet.Cells["A:B"].AutoFitColumns()

    $legendSheet.Cells["A16"].Value = "Provide specific access type information if necessary, e.g."
    $legendSheet.Cells["A16"].Style.Font.Bold = $true

    $legendSheet.Cells["B16"].Value = "Access Type Description"
    $legendSheet.Cells["B16"].Style.Font.Bold = $true

    $accessTypes = [ordered]@{
        "FullControl" = "Allows full control over a file or directory, including reading, writing, changing permissions, and taking ownership."
        "Modify" = "Allows reading, writing, executing, and deleting files and directories."
        "ReadAndExecute" = "Allows reading and executing files."
        "ListDirectory" = "Allows listing the contents of a folder."
        "Read" = "Allows reading data from a file or listing contents of a directory."
        "Write" = "Allows writing data to a file or adding files to a directory."
        "Delete" = "Allows deleting a file or directory."
        "ReadPermissions" = "Allows reading permissions of a file or directory."
        "ChangePermissions" = "Allows changing permissions of a file or directory."
        "TakeOwnership" = "Allows taking ownership of a file or directory."
        "ReadAttributes" = "Allows reading basic attributes of a file or directory."
        "WriteAttributes" = "Allows writing basic attributes of a file or directory."
        "ReadExtendedAttributes" = "Allows reading extended attributes of a file or directory."
        "WriteExtendedAttributes" = "Allows writing extended attributes of a file or directory."
        "ExecuteFile" = "Allows executing a file or traversing a directory."
        "DeleteSubdirectoriesAndFiles" = "Allows deleting subdirectories and files within a directory."
        "Synchronize" = "Allows synchronizing access to a file or directory."
    }

    $row = 17
    foreach ($key in $accessTypes.Keys) {
        $legendSheet.Cells["A$row"].Value = $key
        $legendSheet.Cells["B$row"].Value = $accessTypes[$key]
        $row++
    }
    $row++

    $legendSheet.Cells["A$row"].Value = "Disclaimer"
    $legendSheet.Cells["B$row"].Value = "For more specific information on what the access types do, please refer to online resources such as Google."
    $legendSheet.Cells["A$row:B$row"].Style.Font.Italic = $true

    $legendSheet.Cells["A:B"].AutoFitColumns()

    $excelPackage.Workbook.Worksheets.MoveToStart($legendSheet)

    Close-ExcelPackage $excelPackage
}

function Invoke-OUAudit {
    param (
        [hashtable]$formTag
    )

    try {
        $distinguishedName = $formTag.DistinguishedName
        $folderPath = $formTag.FolderPath
        $folderDepth = $formTag.FolderDepth

        if ($distinguishedName) {
            Write-Host "Attempting to retrieve groups from: $distinguishedName"
        } else {
            Write-Host "`nNo OU was selected or the selected OU does not have a valid distinguished name.`n"
            return
        }

        $LDAPPath = "LDAP://$distinguishedName"
        $entry = New-Object DirectoryServices.DirectoryEntry $LDAPPath
        $searcherOU = New-Object DirectoryServices.DirectorySearcher $entry
        $searcherOU.Filter = "(objectClass=group)"
        $searcherOU.PageSize = 1000
        $groups = $searcherOU.FindAll()
        if ($null -eq $groups) {
            Write-Host "Error: \$groups is null."
            return
        } elseif ($groups.Count -eq 0) {
            Write-Host "Error: No groups found in selected OU."
            return
        }
        Write-Host "Groups retrieval complete."
    }
    catch {
        Write-Host "Error retrieving groups from selected OU: $_"
        return
    }

    try {
        $excelData = @{}
        $processedFolders = @{}
        $allGroupNames = @()
        $groupsWithNoUsers = @()

        foreach ($group in $groups) {
            if ($null -eq $group.Properties.name) {
                Write-Host "Error: \group.Properties.name is null."
                continue
            }

            $groupName = $group.Properties.name[0].ToString()
            $groupEntry = $group.GetDirectoryEntry()

            if ($null -eq $groupEntry) {
                Write-Host "Error: \groupEntry is null."
                continue
            }

            $groupMembers = $groupEntry.psbase.Invoke("Members") | ForEach-Object {
                ($_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)) -replace '^CN=', ''
            }

            $department = $distinguishedName -split ',' | Select-Object -First 1
            $departmentName = $department -split '=' | Select-Object -Last 1

            if (-not $processedFolders.ContainsKey($folderPath)) {
                $processedFolders[$folderPath] = @()
            }

            if ($null -eq $groupMembers) {
                Write-Host "$groupName has no users"
                $groupsWithNoUsers += $groupName
            } else {
                Write-Host "Adding users from $groupName"
            }

            $excelData[$groupName] = @{
                Members = $groupMembers
                Folders = @()
            }

            $processedFolders[$folderPath] += $groupName
            $allGroupNames += $groupName
        }

        $folderAccessResults = Get-FolderAccess -groupNames $allGroupNames -folderPath $folderPath -folderDepth $folderDepth
        foreach ($groupName in $allGroupNames) {
            if ($folderAccessResults.ContainsKey($groupName)) {
                $excelData[$groupName].Folders = $folderAccessResults[$groupName]
            }
        }
    }
    catch {
        Write-Host "Error processing groups or members: $_"
        return
    }

    try {
        if ($excelData.Count -eq 0) {
            Write-Host "No data available for Excel export."
            return
        }
        Write-Host "Exporting to Excel..."

        $OUName = ($distinguishedName -split ',')[0].split('=')[1]
        $departmentName = ($distinguishedName -split ',')[0].split('=')[1]
        $currentDate = Get-Date
        $excelDate = $currentDate.ToString("yyyy-MMM-dd_hh-mmtt")

        $exportsFolder = Join-Path $PSScriptRoot "Exports"
        if (-not (Test-Path -Path $exportsFolder)) {
            New-Item -Path $exportsFolder -ItemType Directory | Out-Null
        }

        $departmentFolder = Join-Path $exportsFolder $departmentName
        if (-not (Test-Path -Path $departmentFolder)) {
            New-Item -Path $departmentFolder -ItemType Directory | Out-Null
        }

        $excelFile = Join-Path $departmentFolder "$OUName-$excelDate.xlsx"
        if (Test-Path -Path $excelFile) {
            Remove-Item -Path $excelFile -Force
        }

        $groupsWithNoUsers | Export-Excel -Path $excelFile -WorksheetName "groups without users"
        $usersHeader = [PSCustomObject]@{ Users = "Users" }
        $worksheetNameMap = @{}
        $worksheetColorMap = @{}
        $sortedExcelData = $excelData.GetEnumerator() | Sort-Object Key
        $sortedExcelData | ForEach-Object {
            if ($_.key.length -gt 31) {
                $sheetName = $_.Key.Substring(0, [Math]::Min($_.Key.Length, 28)) + "..."
                $worksheetNameMap[$sheetName] = $_.Key
            } else {
                $sheetName = $_.Key
            }
            $members = $_.Value.Members | Sort-Object
            $folders = $_.Value.Folders | Sort-Object

            if ($members.Count -eq 0) {
                $worksheetColorMap[$sheetName] = [System.Drawing.Color]::Yellow
            } elseif ($members.Count -ne 0 -and $folders.Count -eq 0) {
                $worksheetColorMap[$sheetName] = [System.Drawing.Color]::LightBlue
            } else {
                $worksheetColorMap[$sheetName] = [System.Drawing.Color]::LightGreen
            }

            $worksheetExists = $false
            if (Test-Path $excelFile) {
                $excelPackage = Open-ExcelPackage -Path $excelFile
                $worksheetExists = ($excelPackage.Workbook.Worksheets.Name -contains $sheetName)
                Close-ExcelPackage $excelPackage -Save:$false
            }

            if ($members.Count -ne 0) {
                $usersHeader | Export-Excel -Path $excelFile -WorksheetName $sheetName -Append:$worksheetExists -StartRow 2
            }
            $members | Export-Excel -Path $excelFile -WorksheetName $sheetName -Append:$worksheetExists -StartRow 3
            $folders | Export-Excel -Path $excelFile -WorksheetName $sheetName -Append:$worksheetExists -StartRow ($members.Count + 4)
        }

        $excelPackage = Open-ExcelPackage -Path $excelFile
        foreach ($worksheet in $excelPackage.Workbook.Worksheets) {
            $lastRow = $worksheet.Dimension.End.Row

            $originalName = if ($worksheetNameMap.ContainsKey($worksheet.Name)) {
                $worksheetNameMap[$worksheet.Name]
            } else {
                $worksheet.Name
            }

            if ($worksheet.Name -eq "groups without users") {
                $worksheet.TabColor = [System.Drawing.Color]::Yellow
                for ($row = 1; $row -le $lastRow; $row++) {
                    $cellA = $worksheet.Cells[$row, 1]
                    $cellB = $worksheet.Cells[$row, 2]
                    $cellA.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $cellA.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
                    $cellB.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $cellB.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
                }
                continue
            } else {
                $backgroundColor = $worksheetColorMap[$worksheet.Name]
                $worksheet.TabColor = $backgroundColor
                $worksheet.Cells["A1"].Value = "$originalName"
                $worksheet.Cells["A1"].Style.Font.Bold = $true
                $worksheet.Cells["A1"].Style.Font.Size = 12
                $worksheet.Cells["A1"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $worksheet.Cells["A1"].Style.Fill.BackgroundColor.SetColor($backgroundColor)
                $worksheet.Cells["B1"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $worksheet.Cells["B1"].Style.Fill.BackgroundColor.SetColor($backgroundColor)
            }

            if ($groupsWithNoUsers.Contains($originalName)) {
                $worksheet.Cells["A2"].Value = "No Users"
            }
            $worksheet.Cells["A2"].Style.Font.Bold = $true
            $worksheet.Cells["A2"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $worksheet.Cells["A2"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
            $worksheet.Cells["B2"].Style.Font.Bold = $true
            $worksheet.Cells["B2"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $worksheet.Cells["B2"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)

            for ($row = 1; $row -le $lastRow; $row++) {
                $cell = $worksheet.Cells[$row, 1]
                if ($cell.Text -eq "Folders") {
                    $cell.Style.Font.Bold = $true
                    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
                    $worksheet.Cells[$row, 2].Style.Font.Bold = $true
                    $worksheet.Cells[$row, 2].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $worksheet.Cells[$row, 2].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
                    break
                }
            }
            $worksheet.Cells.AutoFitColumns()
        }
        Close-ExcelPackage -ExcelPackage $excelPackage
        Add-LegendSheet -excelFile $excelFile -folderPath $folderPath -distinguishedName $distinguishedName
        Write-Host "Exported to $excelFile`n"
    }
    catch {
        Write-Host "Error exporting to Excel: $_"
    }
}

try {
    $rootEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://OU=Dept,DC=USERS,DC=CAMPUS")
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($rootEntry)
    $searcher.Filter = "(objectClass=organizationalUnit)"
    $searcher.SearchScope = [System.DirectoryServices.SearchScope]::OneLevel
    $OUs = $searcher.FindAll()
    Write-Host "Retrieved OUs. Count: $($OUs.Count)`n"
}
catch {
    Write-Host "Error retrieving OUs: $_"
}

try {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "OU Auditer"
    $form.Size = New-Object System.Drawing.Size(400, 430)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.Icon = $icon

    $listView = New-Object System.Windows.Forms.ListView
    $listView.View = [System.Windows.Forms.View]::List
    $listView.Size = New-Object System.Drawing.Size(360, 300)
    $listView.Location = New-Object System.Drawing.Point(10, 10)

    $OUs | Sort-Object { $_.Properties.name[0] } | ForEach-Object {
        $item = New-Object System.Windows.Forms.ListViewItem($_.Properties.name[0].ToString())
        $item.Tag = $_
        $listView.Items.Add($item) | Out-Null
    }
    $form.Controls.Add($listView)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Size = New-Object System.Drawing.Size(360, 20)
    $textBox.Location = New-Object System.Drawing.Point(10, 320)
    $textBox.Text = "Folder Path"
    $textBox.Enabled = $false
    $form.Controls.Add($textBox)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Folder Depth"
    $label.Size = New-Object System.Drawing.Size(70, 20)
    $label.Location = New-Object System.Drawing.Point(10, 355)
    $form.Controls.Add($label)

    $numericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDown.Size = New-Object System.Drawing.Size(60, 20)
    $numericUpDown.Location = New-Object System.Drawing.Point(85, 350)
    $numericUpDown.Minimum = 0
    $numericUpDown.Maximum = 10
    $numericUpDown.Value = 2
    $form.Controls.Add($numericUpDown)

    $depthLabel = New-Object System.Windows.Forms.Label
    $depthLabel.Size = New-Object System.Drawing.Size(100, 20)
    $depthLabel.Location = New-Object System.Drawing.Point(10, 375)
    $depthLabel.ForeColor = [System.Drawing.Color]::Gray
    $depthLabel.Text = ("\x" * $numericUpDown.Value)
    $form.Controls.Add($depthLabel)

    $numericUpDown.Add_ValueChanged({
        $depthLabel.Text = ("\x" * $numericUpDown.Value)
    })

    $runAuditButton = New-Object System.Windows.Forms.Button
    $runAuditButton.Text = "Run Audit"
    $runAuditButton.Size = New-Object System.Drawing.Size(80, 25)
    $runAuditButton.Location = New-Object System.Drawing.Point(290, 350)
    $runAuditButton.Enabled = $false
    $runAuditButton.BackColor = [System.Drawing.Color]::FromArgb(204, 30, 61)
    $runAuditButton.ForeColor = [System.Drawing.Color]::White
    $runAuditButton.Add_Click({
        if ($listView.SelectedItems.Count -eq 0) {
            Write-Host "No item selected inside the OK button click event."
            return
        }
        $selectedTag = $listView.SelectedItems[0].Tag

        if ($selectedTag -and $selectedTag.Properties -and $selectedTag.Properties['distinguishedname'] -and $selectedTag.Properties['distinguishedname'].Count -gt 0) {
            $form.Tag = @{
                DistinguishedName = $selectedTag.Properties['distinguishedname'][0]
                FolderPath = $textBox.Text
                FolderDepth = $numericUpDown.Value
            }
            $form.Hide()
            Invoke-OUAudit -formTag $form.Tag
            $form.Show()
        }
        else {
            Write-Host "Selected item does not have a valid distinguished name."
        }
    })
    $form.Controls.Add($runAuditButton)

    $viewExportsButton = New-Object System.Windows.Forms.Button
    $viewExportsButton.Text = "View Exports"
    $viewExportsButton.Size = New-Object System.Drawing.Size(80, 25)
    $viewExportsButton.Location = New-Object System.Drawing.Point(205, 350)
    $viewExportsButton.BackColor = [System.Drawing.Color]::FromArgb(127, 212, 165)
    $viewExportsButton.ForeColor = [System.Drawing.Color]::White
    $viewExportsButton.Add_Click({
        $exportsFolder = Join-Path $PSScriptRoot "Exports"
        if (Test-Path -Path $exportsFolder) {
            Start-Process "explorer.exe" -ArgumentList $exportsFolder
        } else {
            [System.Windows.Forms.MessageBox]::Show("No Exports folder found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $form.Controls.Add($viewExportsButton)

    $listView.Add_SelectedIndexChanged({
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedTag = $listView.SelectedItems[0].Tag
            if ($selectedTag -and $selectedTag.Properties -and $selectedTag.Properties['distinguishedname'] -and $selectedTag.Properties['distinguishedname'].Count -gt 0) {
                $department = $selectedTag.Properties['distinguishedname'][0] -split ',' | Select-Object -First 1
                $departmentName = $department -split '=' | Select-Object -Last 1
                $textBox.Text = "\\catfiles.users.campus\workarea$\" + $departmentName
                $textBox.Enabled = $true
                $runAuditButton.Enabled = $true
            }
        }
    })

    $form.showDialog() | Out-Null
}
catch {
    Write-Host "Error with GUI or OU selection: $_"
}