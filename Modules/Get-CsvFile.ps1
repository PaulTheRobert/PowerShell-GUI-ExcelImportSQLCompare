function Get-CSVFile {

    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline)]
        [string]$path
    )
    Process{
      
    $csv = Import-Csv $path

    return $csv

    }    
}
