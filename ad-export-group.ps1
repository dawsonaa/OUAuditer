# Load necessary assemblies

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.DirectoryServices

# Ensure the ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "The ImportExcel module is not installed. Running Install-Module command."
    $dialogResult = [System.Windows.Forms.MessageBox]::Show("ImportExcel module is not installed, Attempting to install. Please type 'A' in powershell to install.", "Required Module not installed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    Install-Module -Name ImportExcel -Scope CurrentUser
    #exit
}

Write-Host "Starting the script..."

try {
    # Retrieve top-level OUs from the 'Dept' OU in AD
    $rootEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://OU=Dept,DC=USERS,DC=CAMPUS")
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($rootEntry)
    $searcher.Filter = "(objectClass=organizationalUnit)"
    $searcher.SearchScope = [System.DirectoryServices.SearchScope]::OneLevel
    $OUs = $searcher.FindAll()
    Write-Host "Retrieved OUs. Count: $($OUs.Count)"
    
    <# Revised diagnostic for properties:
if ($OUs -and $OUs.Count -gt 0) {
    $firstOU = $OUs[0]
    
    Write-Host "Properties of the first OU:"
    $firstOU.Properties.PropertyNames | ForEach-Object { Write-Host $_ }

    Write-Host "Distinguished Name of the first OU: $($firstOU.Properties['distinguishedName'][0])"
} else {
    Write-Host "No OUs were retrieved."
}
#>

}
catch {
    Write-Host "Error retrieving OUs: $_"
}


try {
    # GUI to select the OU
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select a Department"
    $form.Size = New-Object System.Drawing.Size(300, 400)
    $form.StartPosition = "CenterScreen"

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
    $okButton.Location = New-Object System.Drawing.Point(10, 320)
    $okButton.Add_Click({
            if ($listView.SelectedItems.Count -eq 0) {
                Write-Host "No item selected inside the OK button click event."
                return
            }
            $selectedTag = $listView.SelectedItems[0].Tag

            # Check for the presence of 'distinguishedname'.
            if ($selectedTag -and $selectedTag.Properties -and $selectedTag.Properties['distinguishedname'] -and $selectedTag.Properties['distinguishedname'].Count -gt 0) {
                #Write-Host "Inside OK Click, selected DN: $($selectedTag.Properties['distinguishedname'][0])"
                $form.Tag = $selectedTag
            }
            else {
                Write-Host "Selected item does not have a valid distinguished name."
                return
            }
    
            <# Additional diagnostics
    if ($form.Tag -and $form.Tag.Properties -and $form.Tag.Properties['distinguishedname'] -and $form.Tag.Properties['distinguishedname'].Count -gt 0) {
        Write-Host "Form.Tag DN after assignment: $($form.Tag.Properties['distinguishedname'][0])"
    } else {
        Write-Host "Error in setting the Form.Tag DN."
        return
    }
    #>

            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
            $form.Dispose()
        })


    $form.Controls.Add($okButton)

    $result = $form.ShowDialog()
    
    <#   Diagnostics
if ($form.Tag -and $form.Tag.Properties -and $form.Tag.Properties['distinguishedname'] -and $form.Tag.Properties['distinguishedname'].Count -gt 0) {
    Write-Host "After closing the form, selected OU DN: $($form.Tag.Properties['distinguishedname'][0])"
    # Previously there was a return statement here, which we should remove, as it prematurely exits the script before the subsequent logic.
} else {
    Write-Host "No OU was selected or the selected OU does not have a valid distinguished name."
    exit
}#>


    <# After the form is closed diagnostics
if ($form.Tag) {
    Write-Host "Form.Tag exists."
    if ($form.Tag.Properties) {
        Write-Host "Form.Tag.Properties exists."
        if ($form.Tag.Properties['distinguishedname']) {
            Write-Host "Form.Tag.Properties['distinguishedname'] exists."
        } else {
            Write-Host "Form.Tag.Properties['distinguishedname'] is missing or empty."
        }
    } else {
        Write-Host "Form.Tag.Properties is missing."
    }
} else {
    Write-Host "Form.Tag is missing."
}
#>

}
catch {
    Write-Host "Error with GUI or OU selection: $_"
}

try {
    Write-Host "Value of distinguishedName: $($form.Tag.Properties["distinguishedName"])"
    $distinguishedName = $form.Tag.Properties["distinguishedName"]
    # After user selects an OU...
    if ($form.Tag -and $form.Tag.Properties -and $form.Tag.Properties['distinguishedname']) {
        #Write-Host "SelectedOU"
        $selectedOU = $form.Tag.Properties["distinguishedName"].ToString()
        Write-Host "Attempting to retrieve groups from: $selectedOU"
    }
    else {
        Write-Host "No OU was selected or the selected OU does not have a valid distinguished name."
        exit
    }

    # Fetch groups within the selected OU
    $selectedOUEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$selectedOU")
    <# Development Diagnostics
if ($null -eq $selectedOUEntry) {
    Write-Host "Error: \$selectedOUEntry is null."
}#>

    $LDAPPath = "LDAP://$($form.Tag.Properties['distinguishedname'])"
    $entry = New-Object DirectoryServices.DirectoryEntry $LDAPPath
    $searcherOU = New-Object DirectoryServices.DirectorySearcher $entry

    if ($null -eq $searcherOU) {
        Write-Host "Error: \$searcherOU is null."
    }

    Write-Host "Setting up the searcher object..."
    $searcherOU.Filter = "(objectClass=group)"
    $searcherOU.PageSize = 1000

    <#
Write-Host "SearchRoot: $($searcherOU.SearchRoot.Path)"
Write-Host "Filter: $($searcherOU.Filter)"
Write-Host "PageSize: $($searcherOU.PageSize)"
#>

    Write-Host "Trying to retrieve all groups..."
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
    # Initialize a hashtable for Excel export
    $excelData = @{}

    # For each group, fetch members and add to hashtable
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
            #Write-Host "Error: \groupMembers is null."
            Write-Host "$groupName has no users"
        }
        else {

            # Add members to hashtable
            Write-Host "Adding users from $groupName"
            $excelData[$groupName] = $groupMembers
        }
    }


}
catch {
    Write-Host "Error processing groups or members: $_"
}

try {
    if ($excelData.Count -eq 0) {
        Write-Host "No data available for Excel export."
        exit
    }
    
    # Assuming you have the distinguished name in the variable: $distinguishedName
    # Extract the OU name (this assumes that the OU name doesn't have any commas in it)
    $OUName = ($distinguishedName -split ',')[0].split('=')[1]

    # Get the current date and format it as MM-DD-YYYY
    $currentDate = Get-Date -Format "MM-dd-yyyy"

    <# Get the default desktop path
$defaultDesktop = "$env:USERPROFILE\Desktop"

# Check if OneDrive desktop redirection is in place
if (Test-Path "$env:USERPROFILE\OneDrive - YourCompanyName\Desktop") {
    $desktopLocation = "$env:USERPROFILE\OneDrive - Kansas State University\Desktop"
} else {
    $desktopLocation = $defaultDesktop
}
#>

    # Create SaveFileDialog object
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop') # Start at the user's desktop
    $saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx" # Filter for .xlsx files
    $saveFileDialog.FileName = "$OUName-GroupExport-$currentDate.xlsx" # Set default file name

    $saveFileOpen = $saveFileDialog.ShowDialog()

    if ($saveFileOpen -eq 'OK') {
        $excelFile = $saveFileDialog.FileName
    }

    #$excelFile = "$desktopLocation\$OUName-GroupExport-$currentDate.xlsx"


    $excelData.GetEnumerator() | ForEach-Object {
        $sheetName = $_.Key
        $members = $_.Value

        # Check if the specific worksheet within the Excel file exists
        $worksheetExists = $false
        if (Test-Path $excelFile) {
            $excelPackage = Open-ExcelPackage -Path $excelFile
            $worksheetExists = ($excelPackage.Workbook.Worksheets.Name -contains $sheetName)
            Close-ExcelPackage $excelPackage -Save:$false
        }

        if ($worksheetExists) {
            # Export to the Excel file and append to existing worksheet
            $members | Export-Excel -Path $excelFile -WorksheetName $sheetName -Append
        }
        else {
            # Export to the Excel file without appending (creates a new worksheet)
            $members | Export-Excel -Path $excelFile -WorksheetName $sheetName
        }
    }

    Write-Host "Exported to $excelFile"
    Write-Host ""
    Read-Host "Press Enter to close the window..."
}
catch {
    Write-Host "Error exporting to Excel: $_"
}
