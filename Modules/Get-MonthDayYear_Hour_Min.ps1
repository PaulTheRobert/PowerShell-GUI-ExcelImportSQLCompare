<#
 .Synopsis
  Displays a form for importng a hospital death list and reconciling it against iTx.
#>

function Get-MonthDayYear_Hour_Min{
[cmdletbinding()]Param()

    #Add the date and time to the file name so that it is always unique - 
    $now = Get-Date
    #prep to append the date and time to the file name the way I wanted
    #the wierd syntax here is regex to make sure I get leading 0s on date parts when a single digit
    $now = '{0:d2}' -f $($now.Month) + '{0:d2}' -f $($now.Day) + $($now.Year) +'_'+ '{0:d2}' -f $($now.Hour) +'_'+ '{0:d2}' -f $($now.Minute)

    return $now
}