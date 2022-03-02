<#
 .Synopsis
 Returns a list of all the hospitals for that OPO

#>
function Get-HospitalList{
[cmdletbinding()]Param(
    $OPO,
    $StartDate,
    $EndDate
)
    $HospitalList = @()

    $ConnectionString = "Data Source=DCIDS-SQL-PROD1;Integrated Security=True;ApplicationIntent=ReadOnly"

    $Qry2 = "Select distinct ReferringOrganizationName From [DCIDSDW].[Dim].[Patient] Where OPO = '$($OPO)' Order By ReferringOrganizationName"

    # $Qry = "USE [DCIDSDW] EXEC HS_HospitalScorecard_Hospital `"$($StartDate)`", `"$($EndDate)`", `"$($OPO)`" "
    $SqlResult = Invoke-Sqlcmd -Query $Qry2 -ConnectionString $ConnectionString 
    write-Host "Get-HospitalList for: $($OPO)"
    write-Host "$($Qry2)"
    return $SqlResult.ReferringOrganizationName
}