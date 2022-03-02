<#
 .Synopsis
 Gets the Xls File

#>

function Get-XlsFile{
[cmdletbinding()]Param(
    $XlsPath,
    $InProcess
) 
    

    write-host "Get-XlsFile: $($XlsPath)"

    $ExcelFile = New-Object -ComObject Excel.Application

    $Workbook = $ExcelFile.Workbooks.Open($XlsPath)

    # we are taking the first tab in the workbook, wookbook should probably have only one tab.

    # TODO you could check the tab count and alert the user if tabs > 1 - I this could be more functional in the future.
    If($Workbook.Worksheets.Count > 1){
        write-host "Workbook should only contain one Tab!"
        }

    $Worksheet = $Workbook.sheets.item(1)

    <# here we could gather some meta data about the selected file for logging purposes 
        -file name, date, user, number of worksheets in the file, name of worksheets in the file, number of columns in each worksheet, name of the columns in each work sheet, how many rows were in each data set
    #>
    
    #Add the date and time to the file name so that it is always unique - 
   $now = Get-MonthDayYear_Hour_Min

    #I want to save this as a csv file.
    #$name = ($Workbook.Name).Replace("xls","csv")

    #instead of replace, add .csv extension at end
    $name = ($Workbook.Name).Replace("xls","")


    
    #save the worksheet as a csv (6) in the $InProcess dir
    $worksheet.SaveAs("$($InProcess)$name_$now.csv",6)

    # make sure you close the workbook at the end
    $ExcelFile.Workbooks.Close()

   
}