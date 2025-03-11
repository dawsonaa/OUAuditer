# Author: Dawson Adams (dawsonaa@ksu.edu)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.DirectoryServices
Add-Type -AssemblyName System.Drawing

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "The ImportExcel module is not installed. Running Install-Module command."
    $dialogResult = [System.Windows.Forms.MessageBox]::Show("ImportExcel module is not installed, Attempting to install. Please type 'A' in powershell to install.", "Required Module not installed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    Install-Module -Name ImportExcel -Scope CurrentUser
}

Write-Host "Starting the script..."

function Get-FolderAccess {
    param (
        [string[]]$groupNames,
        [string]$folderPath,
        [int]$maxDepth = 2
    )
    $startTime = [DateTime]::Now.Ticks

    Write-Host "Getting folder access for groups: $($groupNames -join ', ')`nFolder Path: $folderPath`nMax Depth: $maxDepth`n"

    $accessList = @{}

    $rootAcl = Get-Acl -Path $folderPath
    $rootGroups = $rootAcl.Access | ForEach-Object { $_.IdentityReference.Value }

    foreach ($groupName in $groupNames) {
        $accessList[$groupName] = @()
        if ($rootGroups -contains $groupName) {
            $accessList[$groupName] += [PSCustomObject]@{ Folder = $folderPath; AccessType = $_.FileSystemRights }
        }
    }

    $folders = Get-ChildItem -Path $folderPath -Directory -Recurse -Depth $maxDepth

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

        $folderGroups = $acl.Access | ForEach-Object { $_.IdentityReference.Value }
        $parentGroups = $parentAcl.Access | ForEach-Object { $_.IdentityReference.Value }

        if ($folderGroups -ne $parentGroups) {
            foreach ($groupName in $groupNames) {
                foreach ($access in $acl.Access) {
                    if ($access.IdentityReference -like "*$groupName*") {
                        $accessList[$groupName] += [PSCustomObject]@{ Folder = $folder.FullName; AccessType = $access.FileSystemRights }
                        break
                    }
                }
            }
        }
    }
    $endTime = [DateTime]::Now.Ticks
    Write-Host ("Time taken to get folder access: " + (($endTime - $startTime) / 10000000) + " s")
    return $accessList
}

function Add-LegendSheet {
    param (
        [string]$excelFile
    )

    $excelPackage = Open-ExcelPackage -Path $excelFile

    $legendSheet = $excelPackage.Workbook.Worksheets.Add("Legend")

    $legendSheet.Cells["A1"].Value = "Add user or file location to group"
    $legendSheet.Cells["A2"].Value = "Remove user or file location from group"

    $legendSheet.Cells["B1"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $legendSheet.Cells["B1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen)

    $legendSheet.Cells["B2"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $legendSheet.Cells["B2"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Red)

    $legendSheet.Cells["A3"].Value = ""
    $legendSheet.Cells["A4"].Value = "Please provide the full file path, e.g."
    $legendSheet.Cells["A5"].Value = "\\catfiles.users.campus\workarea$\Dept\Folder\Location"

    $legendSheet.Cells["A1:A2"].Style.Font.Bold = $true
    $legendSheet.Cells["A4"].Style.Font.Bold = $true

    $legendSheet.Cells["A:B"].AutoFitColumns()

    $excelPackage.Workbook.Worksheets.MoveToStart($legendSheet)

    Close-ExcelPackage $excelPackage
}

try {
    $rootEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://OU=Dept,DC=USERS,DC=CAMPUS")
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($rootEntry)
    $searcher.Filter = "(objectClass=organizationalUnit)"
    $searcher.SearchScope = [System.DirectoryServices.SearchScope]::OneLevel
    $OUs = $searcher.FindAll()
    Write-Host "Retrieved OUs. Count: $($OUs.Count)"
}
catch {
    Write-Host "Error retrieving OUs: $_"
}

try {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select a Department"
    $form.Size = New-Object System.Drawing.Size(300, 420)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $listView = New-Object System.Windows.Forms.ListView
    $listView.View = [System.Windows.Forms.View]::List
    $listView.Size = New-Object System.Drawing.Size(260, 300)
    $listView.Location = New-Object System.Drawing.Point(10, 10)

    $OUs | Sort-Object { $_.Properties.name[0] } | ForEach-Object {
        $item = New-Object System.Windows.Forms.ListViewItem($_.Properties.name[0].ToString())
        $item.Tag = $_
        $listView.Items.Add($item) | Out-Null
    }
    $form.Controls.Add($listView)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(10, 350)
    $okButton.Enabled = $false
    $okButton.Add_Click({
        if ($listView.SelectedItems.Count -eq 0) {
            Write-Host "No item selected inside the OK button click event."
            return
        }
        $selectedTag = $listView.SelectedItems[0].Tag

        if ($selectedTag -and $selectedTag.Properties -and $selectedTag.Properties['distinguishedname'] -and $selectedTag.Properties['distinguishedname'].Count -gt 0) {
            $form.Tag = $selectedTag
        }
        else {
            Write-Host "Selected item does not have a valid distinguished name."
            return
        }

        $form.Tag = @{
            DistinguishedName = $selectedTag.Properties['distinguishedname'][0]
            FolderPath = $textBox.Text
            MaxDepth = $numericUpDown.Value
        }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
        $form.Dispose()
    })
    $form.Controls.Add($okButton)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Size = New-Object System.Drawing.Size(260, 20)
    $textBox.Location = New-Object System.Drawing.Point(10, 320)
    $textBox.Text = "Department Root Folder Path"
    $textBox.Enabled = $false
    $form.Controls.Add($textBox)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Max Folder Depth:"
    $label.Size = New-Object System.Drawing.Size(100, 20)
    $label.Location = New-Object System.Drawing.Point(130, 350)
    $form.Controls.Add($label)

    $numericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDown.Size = New-Object System.Drawing.Size(40, 20)
    $numericUpDown.Location = New-Object System.Drawing.Point(230, 350)
    $numericUpDown.Minimum = 1
    $numericUpDown.Maximum = 10
    $numericUpDown.Value = 2
    $form.Controls.Add($numericUpDown)

    $listView.Add_SelectedIndexChanged({
        if ($listView.SelectedItems.Count -gt 0) {
            $selectedTag = $listView.SelectedItems[0].Tag
            if ($selectedTag -and $selectedTag.Properties -and $selectedTag.Properties['distinguishedname'] -and $selectedTag.Properties['distinguishedname'].Count -gt 0) {
                $department = $selectedTag.Properties['distinguishedname'][0] -split ',' | Select-Object -First 1
                $departmentName = $department -split '=' | Select-Object -Last 1
                $textBox.Text = "\\catfiles.users.campus\workarea$\" + $departmentName
                $textBox.Enabled = $true
                $okButton.Enabled = $true
            }
        }
    })

    $result = $form.ShowDialog()
}
catch {
    Write-Host "Error with GUI or OU selection: $_"
}

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    try {
        $formTag = $form.Tag
        $distinguishedName = $formTag.DistinguishedName
        $folderPath = $formTag.FolderPath
        $maxDepth = $formTag.MaxDepth

        if ($distinguishedName) {
            Write-Host "Attempting to retrieve groups from: $distinguishedName"
        }
        else {
            Write-Host "No OU was selected or the selected OU does not have a valid distinguished name."
            exit
        }

        $selectedOUEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$distinguishedName")
        $LDAPPath = "LDAP://$distinguishedName"
        $entry = New-Object DirectoryServices.DirectoryEntry $LDAPPath
        $searcherOU = New-Object DirectoryServices.DirectorySearcher $entry

        $searcherOU.Filter = "(objectClass=group)"
        $searcherOU.PageSize = 1000

        $groups = $searcherOU.FindAll()
        Write-Host "Groups retrieval complete."

        if ($null -eq $groups) {
            Write-Host "Error: \$groups is null."
        }
        elseif ($groups.Count -eq 0) {
            Write-Host "Error: No groups found in selected OU."
        }
    }
    catch {
        Write-Host "Error retrieving groups from selected OU: $_"
        exit
    }

    try {
        $excelData = @{}
        $processedFolders = @{}
        $allGroupNames = @()
        $groupsWithNoUsers = @()

        foreach ($group in $groups) {
            if ($null -eq $group.Properties.name) {
                Write-Host "Error: \group.Properties.name is null."
            }

            $groupName = $group.Properties.name[0].ToString()
            $groupEntry = $group.GetDirectoryEntry()

            if ($null -eq $groupEntry) {
                Write-Host "Error: \groupEntry is null."
            }

            $groupMembers = $groupEntry.psbase.Invoke("Members") | ForEach-Object {
                ($_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)) -replace '^CN=', ''
            }

            if ($null -eq $groupMembers) {
                Write-Host "$groupName has no users"
                $groupsWithNoUsers += $groupName
            }
            else {
                $department = $distinguishedName -split ',' | Select-Object -First 1
                $departmentName = $department -split '=' | Select-Object -Last 1

                if (-not $processedFolders.ContainsKey($folderPath)) {
                    $processedFolders[$folderPath] = @()
                }

                Write-Host "Adding users from $groupName"
                $excelData[$groupName] = @{
                    Members = $groupMembers
                    Folders = @()
                }

                $processedFolders[$folderPath] += $groupName
                $allGroupNames += $groupName
            }
        }

        $folderAccessResults = Get-FolderAccess -groupNames $allGroupNames -folderPath $folderPath -maxDepth $maxDepth

        foreach ($groupName in $allGroupNames) {
            if ($folderAccessResults.ContainsKey($groupName)) {
                $excelData[$groupName].Folders = $folderAccessResults[$groupName]
            }
        }
    }
    catch {
        Write-Host "Error processing groups or members: $_"
        exit
    }

    try {
        if ($excelData.Count -eq 0) {
            Write-Host "No data available for Excel export."
            exit
        }
        Write-Host "Exporting to Excel..."
        $OUName = ($distinguishedName -split ',')[0].split('=')[1]
        $currentDate = Get-Date -Format "MM-dd-yyyy"
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        $saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx"
        $saveFileDialog.FileName = "$OUName-OUAudit-$currentDate.xlsx"

        $saveFileOpen = $saveFileDialog.ShowDialog()

        if ($saveFileOpen -eq 'OK') {
            $excelFile = $saveFileDialog.FileName
        }

        $sortedExcelData = $excelData.GetEnumerator() | Sort-Object Key

        $groupsWithNoUsers | Export-Excel -Path $excelFile -WorksheetName "groups without users"

        $sortedExcelData | ForEach-Object {
            $sheetName = $_.Key
            $members = $_.Value.Members | Sort-Object
            $folders = $_.Value.Folders | Sort-Object

            $worksheetExists = $false
            if (Test-Path $excelFile) {
                $excelPackage = Open-ExcelPackage -Path $excelFile
                $worksheetExists = ($excelPackage.Workbook.Worksheets.Name -contains $sheetName)
                Close-ExcelPackage $excelPackage -Save:$false
            }

            if ($worksheetExists) {
                $members | Export-Excel -Path $excelFile -WorksheetName $sheetName -Append
                $folders | ForEach-Object {
                    [PSCustomObject]@{ Folder = $_.Folder; 'Access Type' = $_.AccessType } | Export-Excel -Path $excelFile -WorksheetName $sheetName -Append -StartRow ($members.Count + 2)
                }
            }
            else {
                $members | Export-Excel -Path $excelFile -WorksheetName $sheetName
                $folders | ForEach-Object {
                    [PSCustomObject]@{ Folder = $_.Folder; 'Access Type' = $_.AccessType } | Export-Excel -Path $excelFile -WorksheetName $sheetName -StartRow ($members.Count + 2)
                }
            }
        }

        $excelPackage = Open-ExcelPackage -Path $excelFile
        foreach ($worksheet in $excelPackage.Workbook.Worksheets) {
            $worksheet.Cells.AutoFitColumns()
        }
        Close-ExcelPackage -ExcelPackage $excelPackage

        Add-LegendSheet -excelFile $excelFile

        Write-Host "Exported to $excelFile"
        Write-Host ""
        Read-Host "Press Enter to close the window..."
    }
    catch {
        Write-Host "Error exporting to Excel: $_"
    }
}