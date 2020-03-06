#
                                            # --- START GUI INITIALIZATION --- #
#
#
# Load required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
 
# --- START FUNCTIONS --- #

# Second function. Requires a CSV file extension
function open_CSV_File{
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = "C:\";
    $OpenFileDialog.Filter = "csv files (*.csv)|*.csv"
    
    # If we click open instead of canecl than the box will show and populate our text box with what we selected in the dialog
    if ($OpenFileDialog.ShowDialog() -eq "OK"){
        $textbox_BrowseForCSV.Text = $OpenFileDialog.FileName 
    }
}
#
#
                                                # --- END FUNCTIONS --- #
#------------------------------------------------------------------------------------------
                                                # --- START GUI WINDOW --- # 

#
# Window default settings
# Drawing form and controls
$Form_MachineInfo = New-Object System.Windows.Forms.Form
    $Form_MachineInfo.Text = "Check Machine List in AD"
    $Form_MachineInfo.Size = New-Object System.Drawing.Size(600,300)
    $Form_MachineInfo.FormBorderStyle = "FixedDialog"
    $Form_MachineInfo.TopMost = $false
    $Form_MachineInfo.MaximizeBox = $false
    $Form_MachineInfo.MinimizeBox = $true
    $Form_MachineInfo.ControlBox = $true
    $Form_MachineInfo.StartPosition = "CenterScreen"
    $Form_MachineInfo.Font = "Aileron Bold"
    $ClickFont = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Regular)
    $ClickedFont = [System.Drawing.Font]::new('Aileron Bold', 18, [System.Drawing.FontStyle]::Italic)
 # How far below the window something is
$Y = 60
$Yspacer = 32


# Global gui variables
$LabelSize_universal = New-Object System.Drawing.Size(160,20)
$buttonSize_universal = New-Object System.Drawing.Size(96,23)
$TextBoxSize_universal = New-Object System.Drawing.Size(240,20)
$TextAlign_universal = "MiddleRight"

$Y += $Yspacer


# -- Specify CSV file from disk location -- #
# adding a label
$LabelLabel_BrowseForCSV = New-Object System.Windows.Forms.Label
    $LabelLabel_BrowseForCSV.Location = New-Object System.Drawing.Size(8,$Y)
    $LabelLabel_BrowseForCSV.Size = $LabelSize_universal
    $LabelLabel_BrowseForCSV.TextAlign = $TextAlign_universal
    $LabelLabel_BrowseForCSV.Text = "Open CSV Location"
        $Form_MachineInfo.Controls.Add($LabelLabel_BrowseForCSV)
 
# Adding a textbox for the path to show
$textbox_BrowseForCSV = New-Object System.Windows.Forms.TextBox
    $textbox_BrowseForCSV.Location = New-Object System.Drawing.Size(192,$Y)
    $textbox_BrowseForCSV.Size = $TextBoxSize_universal
        $Form_MachineInfo.Controls.Add($textbox_BrowseForCSV)

# add a button
$button_BrowseForCSV = New-Object System.Windows.Forms.Button
    $button_BrowseForCSV.Location = New-Object System.Drawing.Size(448,$Y)
    $button_BrowseForCSV.Size = $buttonSize_universal
    $button_BrowseForCSV.Text = "Browse"
    $button_BrowseForCSV.Add_Click({open_CSV_File})
        $Form_MachineInfo.Controls.Add($button_BrowseForCSV)
# -- End Open File from Disk Location -- #
#
#
                                             # --- END GUI WORK --- #

# -----------------------------------------------------------------------------------------------------------------------------------

                                            #  --- START MAIN SCRIPT --- #
#
#
function script_Main{
    #if the csv path is not empty do the following:
    if ($textbox_BrowseForCSV.Text -ne ""){
        #Gets the list of computer names from the users chosen path (line 31)
        $names = Get-ADComputer -Filter * -Properties IPv4Address | Select-Object -Property Name,IPv4Address | where-object {$_.ipv4address -ne $null}
         

        #Exports the entire file to csv
        $names| Export-Csv -Path $textbox_BrowseForCSV.Text -NoTypeInformation   
    }
    else {
        Write-Host "          Please choose a file"
    }

    #Script ends with opening the CSV file automatically

    Invoke-Item $textbox_BrowseForCSV.Text

    # Gets the parent directory path of the file
    $folderDirectory = Split-Path -Path "$($textbox_BrowseForCSV.Text)"

    # Sets location to the directory that all the files are contained in
    Set-Location "$folderDirectory"
}

                                            # --- END script function --- #

# -----------------------------------------------------------------------------------------------------------------------------------
                                     # --- Run script and add more gui elements --- #
                                            
# This gui element needed to be down here since it uses the script_Main function
# --- Make a run button for script --- #

# -- Specify text file from disk location -- #
# adding a label for the text box and button
$Label_RunScript = New-Object System.Windows.Forms.Label
    $Label_RunScript.Location = New-Object System.Drawing.Size(18,$Y)
    $Label_RunScript.Size = $LabelSize_universal
    $Label_RunScript.TextAlign = $TextAlign_universal
    $Label_RunScript.Text = "Run"
        $Form_MachineInfo.Controls.Add($Label_RunScript)

# adds a button so we for 'Run Script'
$button_RunScript = New-Object System.Windows.Forms.Button
    $button_RunScript.Location = New-Object System.Drawing.Size(250,150)
    $button_RunScript.Size = New-Object System.Drawing.Size(96,50)
    $button_RunScript.Text = "Run Script"
    $button_RunScript.Add_Click({script_Main})
        $Form_MachineInfo.Controls.Add($button_RunScript)



# This basically adds the window to the scene and displays it on the users screen.
$Form_MachineInfo.Add_Shown({$Form_MachineInfo.Activate()})
[void] $Form_MachineInfo.ShowDialog()



