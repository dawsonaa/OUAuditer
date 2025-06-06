# <img src="icon.ico" alt="Icon" height="32" style="vertical-align: middle;"> <span style="vertical-align: middle;">OU Auditer</span>  <img src="icon.ico" alt="Icon" height="32" style="vertical-align: middle;">

## By Dawson Adams (dawsonaa@ksu.edu) at Kansas State University </span>

## Overview
This PowerShell script exports Active Directory (AD) group memberships for a specified Organizational Unit (OU). Additionally, it retrieves W drive folder access permissions associated with these groups and their members. This audit ensures that groups provide the intended access and verifies group membership accuracy.

## GUI
<img src="images/gui.png" alt="GUI Screenshot">

## Legend Sheet
<img src="images/legend-sheet.png" alt="Legend Sheet Screenshot">

## Summary Sheet
<img src="images/summary-sheet.png" alt="Summary Sheet Screenshot">

## Group Sheet
<img src="images/group-sheet.png" alt="Group Sheet Screenshot">

## Features
- Extracts AD group memberships for a selected OU.
- Retrieves associated W drive folder access permissions.
- Exports organized Excel reports, including:
    - Group memberships.
    - Folder access permissions.
    - A legend sheet for easy reference.
- Allows quick access to exported reports via the **"View Exports"** button in the GUI.

## Prerequisites
- PowerShell 5.1 or later.
- Active Directory module for PowerShell.
- Sufficient permissions to query AD and access W drive permissions.

## Usage
1. **Open the application** by clicking the icon on the desktop or from the start menu  
   <img src="images/desktop-icon.png" alt="Desktop Icon" height="200"><img src="images/startmenu-icon.png" alt="Start Menu Icon" height="200">.
2. **Specify the target OU** and the target network folder location when prompted.
3. **Review the generated report** by pressing the **"View Exports"** button on the GUI to open your exports folder containing the organized Excel files.