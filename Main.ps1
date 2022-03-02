$Dir = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)

#import the modules
foreach($Module in Get-ChildItem $Dir\Modules){
    Import-Module $Module.PSPath
}

# add in a Year month day file nam

#TODO - Uptop we should check for the sub directory structure that we need.
# Create any missing dirictories
    md -force $Dir\WorkingFiles\InProcess
    md -force $Dir\WorkingFiles\Complete
    md -force $Dir\WorkingFiles\ReportOut

#TODO Clean up any stale files that dont belong.

# $InProcess = "C:\Users\patrick.spearman\Desktop\DRR-master\WorkingFiles\InProcess\" 
# $Complete = "C:\Users\patrick.spearman\Desktop\DRR-master\WorkingFiles\Complete\" 

$InProcess = "$Dir\WorkingFiles\InProcess\"
$Complete = "$Dir\WorkingFiles\Complete\"           #these are copies of the input file
$ReportOut = "$Dir\WorkingFiles\ReportOut\"         #these are the output files

# NOTE: filename might not be a very clear variabe name. Consider revising that.
# I am returning a PSCustomObject instead of a single variabe from the form now
#$fileName = Get-Form

#get the parameters back from the Get-Form module. These are the user inputs
$FormParameter = Get-Form

# parse the custom object that is returned from Get-Form
$InputFileName = $FormParameter.InputFileName
$OPO = $FormParameter.OPO
$Hospital = $FormParameter.Hospital
$StartDate = $FormParameter.StartDate
$EndDate = $FormParameter.EndDate

#get only the filename of the path
$InputFileNameParts = $InputFileName.Split("\")
$fileName = $InputFileNameParts[$InputFileNameParts.Count-1]


#TODO this form inputs could also be logged!
Write-Host $FormParameter

#take the first tab of the xls file and saves it as a csv file in the inProcess directory
Get-XlsFile $InputFileName $InProcess

#load the contents of the file into the CSV variable
#might want to check to make sure there is only one file in process at any given time 
foreach($item in Get-ChildItem $InProcess -File -Recurse){
    $csv = Get-CSVFile $item.FullName
} 

#TODO (Paul Davis 01/25/21)  
# Processing I am not sure if I want to seperate this into a module or leave it in main
# as I go, I think this is where the challenge scaling this with other hospitals comes in 
# there will be a need to handle a transformation for the lack of uniformity between hospital files

# NOTE: Don't forget this initial prototype was only designed to handle Vandy ('Vanderbilt University Medical Center')

# unless with the timing of the CMS meterics we actuall can get a quality standardized file . . . ? 
# Maybe also expect this to start pushing run data to a Data Base
# Collecting ICD 10 Coded to inform CMS denom potentially? Might have traction? 

#initialize for collection of Hospital Cases from the hospital death list selected by the user
$HospitalCases = @()

#TODO Gari suggested SOMETIMES DEATHE DATE WILL BE NULL these records are being excluded at the moment. 


# 1) Get the SQL data for the given params as a collection
    $iTxReferrals = Get-Referrals $StartDate $EndDate $Hospital

# 2) Comparison Processing           
    # getting a group of blank rows from the CSV - clean that up
    $csv = $csv | Where-Object {$_.MRN -ne ''}

    #iterate through the CSV first and load it into an object
    foreach($hsCase in $csv){
    
    # A) Make a Custom PS object with an element for each column in each data set 
        # i)  - clearly prefix elements so its easy to tell which source the column originated in

        $HospitalCase = [PSCustomObject]@{
            # hs_ for "hospital"      - from the original input de$Reath list sent from the hospital
                'hs_Name'          = $hsCase.Name
                'hs_Pt Previous Name' = $hsCase.'Pt Previous Name'
                'MRN'              = $hsCase.MRN
                'Discharge Unit'   = $hsCase.'Discharge Unit'
                'Age'              = $hsCase.Age
                'Time of Death'    = $hsCase.'Time of Death'
                'Date of Death'    = $hsCase.'Date of 
Death'
                'Attending'        = $hsCase.Attending
                'Funeral Home'     = $hsCase.'Funeral Home'
                'Pt Class'         = $hsCase.'Pt Class'

        # ii) - include additional logical fieldas flags for processing results - Match Referral number when a match 
                #These are processing flags being initialized blank here maybe?                
                #  - maybe a yes/no text field?
                'iTx_MatchFound'   = ''

                #  - for ReferralNumber of the matching case
                # I am trying to parralell the SQL sp syntax here fyi - i thought that would make it most obvious what data I am returning

	            'PAT.[OPO]'                            = '' 
	            'PAT.[ReferralNumber]'                 = '' 
	            'PAT.[ReferralDate]'                   = '' 
	            'PAT.[DeathDate]'                      = '' 
	            'PAT.[DeathType]'                      = '' 
	            'PAT.[ReferringOrganizationName]'      = '' 

	            'PAT.[MRN]'                            = '' 
	            'PAT.[FirstName]'                      = '' 
	            'PAT.[MiddleName]'                     = '' 
	            'PAT.[LastName]'                       = '' 
	            'PAT.[Age]'                            = '' 
	            'PAT.[AgeUnits]'                       = '' 
	            'PAT.[Gender]'                         = '' 
	            'PAT.[RacePrimary]'                    = '' 
        }

            
        # add each object to the collection
        $HospitalCases += $HospitalCase
    }      

    # B) Check the cases sent on the hospital death list against the sql results and ID anything missing in the SQL Results
        # i) loop through each Hospital Case from the hospital death list
        #Test var counters
        $CaseNo = 0
        $MatchCount1 = 0
        $ExecutionDateTime = Get-Date            #for logging purposes that I just made up

        Foreach($HospitalCase in $HospitalCases){
            # a) Now loop through all of the iTx referrals and check for MRN number matches
            
            # Test Var 
            $iTxCasesChecked = 0        #counter
            $isMatch = 0               #flag if matched
            $matchCase = ''             #Referral number if match

            #incriment
            $CaseNo = $CaseNo + 1

            Foreach($iTxReferral in $iTxReferrals){
                # here i am using like to get an aprox match without leading or lagging 0 's hopefully
                # I wonder if i need to check in both directions? <-- ** LOOKOUT ** - thats a wierd thought
                
                #incriment 
                $iTxCasesChecked = $iTxCasesChecked + 1

                # !!** MATCH **!!
                If($HospitalCase.MRN -like "*$($iTxReferral.MRN)*"){
                    #incriment counter
                    $MatchCount1 = $MatchCount1 +1
                    
                    $matchCase = $iTxReferral.ReferralNumber
                    $isMatch   = 1

                    $HospitalCase.iTx_MatchFound   = 1

                    #  - for ReferralNumber of the matching case
                    # I am trying to parralell the SQL sp syntax here fyi - i thought that would make it most obvious what data I am returning
	                
                    $HospitalCase.'PAT.[OPO]'               = $iTxReferral.OPO                
                    $HospitalCase.'PAT.[ReferralNumber]'            = $iTxReferral.ReferralNumber	                
                    $HospitalCase.'PAT.[ReferralDate]'              = $iTxReferral.ReferralDate	                
                    $HospitalCase.'PAT.[DeathDate]'                 = $iTxReferral.DeathDate	                
                    $HospitalCase.'PAT.[DeathType]'                 = $iTxReferral.DeathType	                
                    $HospitalCase.'PAT.[ReferringOrganizationName]' = $iTxReferral.ReferringOrganizationName	                
                    $HospitalCase.'PAT.[MRN]'                       = $iTxReferral.MRN	                
                    $HospitalCase.'PAT.[FirstName]'                 = $iTxReferral.FirstName	                
                    $HospitalCase.'PAT.[MiddleName]'                = $iTxReferral.MiddleName	                
                    $HospitalCase.'PAT.[LastName]'                  = $iTxReferral.LastName	                
                    $HospitalCase.'PAT.[Age]'                       = $iTxReferral.Age	                
                    $HospitalCase.'PAT.[AgeUnits]'                  = $iTxReferral.AgeUnits	                
                    $HospitalCase.'PAT.[Gender]'                    = $iTxReferral.Gender	                
                    $HospitalCase.'PAT.[RacePrimary]'               = $iTxReferral.RacePrimary
                    
                    #exit the loop on match
                    Break
                }
            }
    # C) Check the cases in the result set from iTx and return any that are missing from the death list

            #TODO write in a log file creation by username_date? into the working dir
            #sanity checkpoint - also maybe make some rudimentry logging?
            #Write-Host "CaseNo: $($CaseNo); MatchCount1: $($MatchCount1); isMatch: $($isMatch); matchCase: $($matchCase); HospitalMRN: $($HospitalCase.MRN); iTxMRN: $($iTxReferral.MRN);iTxCasesChecked: $($iTxCasesChecked); userName: $($env:USERNAME); fileName: $($fileName) ; OPO: $($SelectedOpo); Hospital: $($SelectedHospital); StartDate: $($StartDate); EndDate: $($EndDate); ExecutionDateTime: $($ExecutionDateTime)"
            Write-Host "CaseNo: $($CaseNo); MatchCount1: $($MatchCount1); isMatch: $($isMatch); matchCase: $($matchCase); HospitalMRN: $($HospitalCase.MRN); iTxMRN: $($iTxReferral.MRN);iTxCasesChecked: $($iTxCasesChecked);"
}

        
# Clear InProcess
Get-ChildItem $InProcess -File -Recurse | Move-Item -Destination $Complete -Force

$Unique = Get-MonthDayYear_Hour_Min

$fileName = $fileName.Replace(".xls","")
 
$path = "$($ReportOut)$($fileName)_$($Unique).csv"

# Write Output file
$HospitalCases | Export-Csv -Path $path