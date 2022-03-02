<#
 .Synopsis
 Returns a list of all the referrals within the paramerter range

#>
function Get-Referrals{
[cmdletbinding()]Param(
    $StartDate,
    $EndDate,
    $Hospital
) 
    $ReferralList = @()

    # LOOKOUT!! This is connected to DEV for now - I didn't want to just push the SP to PROD while this is still in process
    $ConnectionString = "Data Source=DCIDS-SQL-DEV1;Integrated Security=True;ApplicationIntent=ReadOnly"

    #FYI Start Date and End date in the SQL are being applied to PAT.[DeathDate] <-- **LOOKOUT**
    $Qry = "USE [DCIDSDW] EXEC dbo.[DS_Drr] '$($StartDate)', '$($EndDate)',   '$($Hospital)'"
    
    #Run the query get the results
    $SqlResult = Invoke-Sqlcmd -Query $Qry -ConnectionString $ConnectionString 

    #TODO The result count would probable be a good thing to log
    Write-Host $ResultCount = $SqlResult.Count


    return $SqlResult
}