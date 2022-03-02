<#
 .Synopsis
  Displays a form for importng a hospital death list and reconciling it against iTx.
#>

#TODO
# Can we add some validation to make sure the user populates all of the fields on the form - I dont think it works if any are left blank

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
function Get-Form{
[cmdletbinding()]Param()
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $Global:HospitalList = @()
    $Global:SelectedOpo = 'BLANK'
    $Global:SelectedHospital = 'BLANK'

    Function SelectFile (){
        # TODO set the initial directory to the i: drive path for the DRR team?
        
        # $initialDirectory = "C:\Users\patrick.spearman\Desktop\DRR-master\TestFiles"
        $initialDirectory = "C:\Users\paul.davis\source\repos\DeathRecordReview\TestFiles"

        $filePicker = New-Object System.Windows.Forms.OpenFileDialog
        $filePicker.Title = "Please Select a File"
        $filePicker.InitialDirectory = $initialDirectory
        $filePicker.Filter = "All files (*.*) | *.*"
        $filePicker.FilterIndex = 2
    
        If ($filePicker.ShowDialog() -eq "Cancel"){
            [System.Windows.Forms.MessageBox]::Show("No File Selected. Please select a file !", "Error", 0, 
            [System.Windows.Forms.MessageBoxIcon]::Exclamation)
        }   
        else{
           $return = $filePicker.FileName
           
        }
        $filePicker.Dispose()
        return $return
    }

    Function SelectOpo($StartDate, $EndDate){
        $today     = Get-Date
        $StartDate = "$($today.AddMonths(-6).month)/$($today.AddMonths(-6).day)/$($today.AddMonths(-6).year)"
        $EndDate   = "$($today.month)/$($today.day)/$($today.year)"

        $Global:SelectedOpo = $OpoDropDown.SelectedItem
        write-host "Selected OPO: $($Global:SelectedOpo)"
        $OPO = $Global:SelectedOpo
        $HospitalList = Get-HospitalList $OPO $StartDate $EndDate
        #Write-Host $HospitalList
        Return $HospitalList
    }

    Function SelectHospital(){
        $Global:SelectedHospital = $HospitalDropDown.SelectedItem
        Write-Host "Selected Hospital: $($Global:SelectedHospital)"
    }

 

    #Form Settings
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'DCIDS DRR Tool'
    $form.Size = New-Object System.Drawing.Size(600,480)
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(7,14)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = 0
    $form.MinimizeBox = 0
    $form.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide
    $form.Margin = New-Object System.Windows.Forms.Padding(3,2,3,2)

    #Header for Form
    $formHeader = New-Object System.Windows.Forms.Label
    $formHeader.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 14)
    $formHeader.Location = New-Object System.Drawing.Point(162,9)
    $formHeader.Size = New-Object System.Drawing.Size(199,29)
    $formHeader.Text = 'DCIDS DRR Tool'
    $formHeader.Name = 'DCIDS_DRR_Tool'
    $formHeader.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
    $form.Controls.Add($formHeader)

    #Choose File Label
    $SelectFileLabel = New-Object System.Windows.Forms.Label
    $SelectFileLabel.Location = New-Object System.Drawing.Point(25,63)
    $SelectFileLabel.Size = New-Object System.Drawing.Size(169,17)
    $SelectFileLabel.Text = 'Choose the File to Import:'
    $SelectFileLabel.Name = 'Choose_File_Lbl'
    $form.Controls.Add($SelectFileLabel)

    #Choose File TB
    $fileTxtBox = New-Object System.Windows.Forms.TextBox
    $fileTxtBox.Location = New-Object System.Drawing.Point(28,82)
    $fileTxtBox.Margin = New-Object System.Windows.Forms.Padding(3,2,3,2)
    $fileTxtBox.Size = New-Object System.Drawing.Size(304,22)
    $fileTxtBox.Text = ''
    $fileTxtBox.Name = 'Choose_File_TB'
    $form.Controls.Add($fileTxtBox)

    #OPO Label
    $OPODropDownLabel = New-Object System.Windows.Forms.Label
    $OPODropDownLabel.Location = New-Object System.Drawing.Point(26,132)
    $OPODropDownLabel.Size = New-Object System.Drawing.Size(43,17)
    $OPODropDownLabel.Text = 'OPO:'
    $OPODropDownLabel.Name = 'OPO_Lbl'
    $form.Controls.Add($OPODropDownLabel)

    #OPO DDL
    $OpoDropDown = New-Object System.Windows.Forms.ComboBox
    $OPOList = @("NMDS","SDS","TDS")
    $OpoDropDown.Items.AddRange($OPOList)
    $OpoDropDown.Location = New-Object System.Drawing.Point(28,151)
    $OpoDropDown.Margin = New-Object System.Windows.Forms.Padding(3,2,3,2)
    $OpoDropDown.Size = New-Object System.Drawing.Size(454,24)
    $OpoDropDown.Name = 'OPO_DDL'
    $form.Controls.Add($OpoDropDown)

    #Hospital Label
    $HospitalDropDownLabel = New-Object System.Windows.Forms.Label
    $HospitalDropDownLabel.Location = New-Object System.Drawing.Point(26,198)
    $HospitalDropDownLabel.Size = New-Object System.Drawing.Size(63,17)
    $HospitalDropDownLabel.Text = 'Hospital:'
    $HospitalDropDownLabel.Name = 'Hospital_Lbl'
    $form.Controls.Add($HospitalDropDownLabel)

    #Hospital DDL
    $HospitalDropDown = New-Object System.Windows.Forms.ComboBox
    #   $HospitalDropDown.Items.AddRange($HospitalList)
    $HospitalDropDown.Location = New-Object System.Drawing.Point(28,217)
    $HospitalDropDown.Margin = New-Object System.Windows.Forms.Padding(3,2,3,2)
    $HospitalDropDown.Size = New-Object System.Drawing.Size(454,24)
    $HospitalDropDown.Name = 'Hospital_DDL'
    $form.Controls.Add($HospitalDropDown)

    #Start Date Label
    $StartDateLbl = New-Object System.Windows.Forms.Label
    $StartDateLbl.Location = New-Object System.Drawing.Point(25,273)
    $StartDateLbl.Size = New-Object System.Drawing.Size(76,17)
    $StartDateLbl.Text = 'Start Date'
    $StartDateLbl.Name = 'Start_Date_Lbl'
    $form.Controls.Add($StartDateLbl)

    #Start Date TB
    $StartDateTB = New-Object System.Windows.Forms.DateTimePicker
    $StartDateTB.Location = New-Object System.Drawing.Point(28,292)
    $StartDateTB.Margin = New-Object System.Windows.Forms.Padding(3,2,3,2)
    $StartDateTB.Size = New-Object System.Drawing.Size(189,24)
    #$StartDateTB.Text = ''
    $StartDateTB.Name = 'Start_Date_TB'
    $StartDateTB.Format = [Windows.Forms.DateTimePIckerFormat]::Custom
    $StartDateTB.CustomFormat = 'MM/dd/yyyy'   
    $todayDT = Get-Date -DisplayHint Date
    $firstDayMonthDT = $todayDT.AddDays(-($todayDT.Day - 1))
    $firstDateLastMonth = $firstDayMonthDT.AddMonths(-1)
    $StartDateTB.Value = $firstDateLastMonth
    $form.Controls.Add($StartDateTB)

    #End Date Label
    $EndDateLbl = New-Object System.Windows.Forms.Label
    $EndDateLbl.Location = New-Object System.Drawing.Point(290,273)
    $EndDateLbl.Size = New-Object System.Drawing.Size(71,17)
    $EndDateLbl.Text = 'End Date'
    $EndDateLbl.Name = 'End_Date_Lbl'
    $form.Controls.Add($EndDateLbl)

    #End Date TB
    $EndDateTB = New-Object System.Windows.Forms.DateTimePicker
    $EndDateTB.Location = New-Object System.Drawing.Point(293,292)
    $EndDateTB.Margin = New-Object System.Windows.Forms.Padding(3,2,3,2)
    $EndDateTB.Size = New-Object System.Drawing.Size(189,24)
    #$EndDateTB.Text = ''
    $EndDateTB.Name = 'End_Date_TB'
    $EndDateTB.Format = [Windows.Forms.DateTimePIckerFormat]::Custom
    $EndDateTB.CustomFormat = 'MM/dd/yyyy'   
    $lastDayLastMonth = $firstDayMonthDT.AddDays(-1)
    $EndDateTB.Value =$lastDayLastMonth
    $form.Controls.Add($EndDateTB)

    #OK/Submit Button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(290,345)
    $okButton.Margin = New-Object System.Windows.Forms.Padding(3,2,3,2)
    $okButton.Size = New-Object System.Drawing.Size(86,29)
    $okButton.Text = 'Submit'
    $okButton.Name = 'Submit_Btn'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    #Cancel Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(395,345)
    $cancelButton.Margin = New-Object System.Windows.Forms.Padding(3,2,3,2)
    $cancelButton.Size = New-Object System.Drawing.Size(87,29)
    $cancelButton.Text = 'Cancel'
    $cancelButton.Name = 'Cancel_Btn'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    #Select File Button
    $getFileButton = New-Object System.Windows.Forms.Button
    $getFileButton.Location = New-Object System.Drawing.Point(373,78)
    $getFileButton.Margin = New-Object System.Windows.Forms.Padding(3,2,3,2)
    $getFileButton.Size = New-Object System.Drawing.Size(108,27)
    $getFileButton.Text = 'Select File'
    $getFileButton.Name = 'Select_File_Btn'
    $form.Controls.Add($getFileButton)

    $getFileButton.Add_Click({$fileTxtBox.Text = SelectFile})

    $OPODropDown.Add_SelectedIndexChanged({
        $HospitalList = SelectOpo
        write-host "Hospital List: $($HospitalList)"
        $HospitalDropDown.Items.AddRange($HospitalList)
    })

    #when the selected index of the Hospital Dropdown is changed, the SelectHospital function (up top ~ ln. 61 ) is called
    $HospitalDropDown.Add_SelectedIndexChanged({SelectHospital})

    $form.Topmost = $true

    [void]$form.ShowDialog()

    $fileName = $fileTxtBox.Text

    $SelectedStartDate = $StartDateTB.Text
    $SelectedEndDate = $EndDateTB.Text

    write-host "OPO: $($SelectedOpo)"

    #I need to get more than one variable back from the form, i think I will use a custom object and let the calling (Main) parse it
    $FormParameter = [PSCustomObject]@{
        InputFileName = $fileName
        OPO =  $SelectedOpo
        Hospital = $SelectedHospital
        StartDate = $SelectedStartDate
        EndDate =  $SelectedEndDate
    }

    #return $fileName 
    return $FormParameter

    $form.Dispose()
}
#Export-ModuleMember -Function Get-Form -Alias *
