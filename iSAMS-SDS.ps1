# Script to create CSV files for SDS (v2) based off data held in iSAMS MIS database.
# Data exported via iSAMS API
#
#
# Katy Nicholson, May 2021
# https://katystech.blog/
# v 1.0.5
#

[CmdletBinding(DefaultParameterSetName="Main")]
param([switch][Parameter(ParameterSetName="ViewHelp")]$Help,
[switch][Parameter(ParameterSetName="Main")]$SkipDownload,
[string][Parameter(Mandatory=$true, ParameterSetName="Main")]$CSVPath,
[switch][Parameter(ParameterSetName="ViewSubject")]$ViewSubjects,
[string[]][Parameter(ParameterSetName="Main")]$ExcludedSubjects,
[string[]][Parameter(ParameterSetName="Main")]$IncludedSubjects,
[string[]][Parameter(ParameterSetName="Main")]$YearGroups,
[string][Parameter(ParameterSetName="Main")]$Suffix,
[string[]][Parameter(ParameterSetName="Main")]$ExtraTeachers
)

# Fill out variables:
$SchoolName = "SchoolName"		                          #Used in SDS CSV as school name
$DFENo = "8889999"				                          #Used in SDS CSVs as school ID, can be anything you like but I tend to use the school's DFE number
$iSAMSHost = "isams.school.net"	                          #Hostname of your iSAMS server
$iSAMSAPIKey = "8B87C910-FFFF-DDDD-AAAA-E0FAABCDEF1C"	  #API Key (looks like a GUID)


if ($Help) {
    Write-Host
    Write-Host "Example: .\iSAMS-SDS.ps1 -CSVPath C:\SDS-CSV [-SkipDownload -ExcludedSubjects @(44,16,17) -YearGroups @(7,8,9,10,11,12,13) -Suffix '/2020/21' -ExtraTeachers @('DeputyHead@school.com')]"
    Write-Host "All parameters except CSVPath are optional"
    Write-Host "ExcludedSubjects is an array of Subject IDs you wish to exclude. Run iSAMS-SDS.ps1 -ViewSubjects to see a list of ID/Names"
    Write-Host "IncludedSubjects is an array of Subject IDs you wish to include. Don't use this at the same time as ExcludedSubjects"
    Write-Host "YearGroups selects the year groups you want data for. No value denotes years 0-13."
    Write-Host "SkipDownload skips downloading the iSAMS data, if you've already got it recently, more useful when testing"
    Write-Host "Suffix is added to each team name"
    Write-Host "ExtraTeachers is an array of the UPN of any teachers you want added to every team, e.g. SLT"
    return
}

if (!$YearGroups) {
    $YearGroups = @(0,1,2,3,4,5,6,7,8,9,10,11,12,13)
}

$xmlFile = $env:TEMP + "\sds_isams_data.xml"

if (-not($SkipDownload)) {
    Invoke-WebRequest -Uri "https://$iSAMSHost/api/batch/1.0/xml.ashx?apiKey=$iSAMSAPIKey" -OutFile $xmlFile
}
[xml]$isamsData = Get-Content -Path $xmlFile


if ($ViewSubjects) {
    foreach ($department in $isamsData.iSAMS.TeachingManager.Departments.Department) {
        foreach ($subject in $department.Subjects.Subject) {
            Write-Output ([PSCustomObject]@{
                "Subject ID" = $subject.Id
                "Code" = $subject.Code
                "Name" =$subject.Name 
            })
        }
    }
    return
}

if ($IncludedSubjects) {
    foreach ($department in $isamsData.iSAMS.TeachingManager.Departments.Department) {
        foreach ($subject in $department.Subjects.Subject) {
            if ($IncludedSubjects -notcontains $subject.Id) {
                $ExcludedSubjects += $subject.Id
            }
        }
    }
   
}



$Errors = @()

# Remove old files if they exist
foreach ($file in @("users","orgs","classes","enrollments")) {
    if (Test-Path("$CSVPath\$file.csv")) {
        Remove-Item "$CSVPath\$file.csv"
    }
}


#Set up the new files
Add-Content -Path "$CSVPath\orgs.csv" -Value "sourcedId,name,type"
Add-Content -Path "$CSVPath\classes.csv" -Value "sourcedId,orgSourcedId,title"
Add-Content -Path "$CSVPath\enrollments.csv" -Value "classSourcedId,userSourcedId,role"
Add-Content -Path "$CSVPath\users.csv" -Value "sourcedId,orgSourcedIds,username,role,familyName,givenName,password,grade"
Add-Content -Path "$CSVPath\orgs.csv" -Value "$DFENo,$SchoolName,school"


#Process pupils
foreach ($entry in $isamsData.iSAMS.PupilManager.CurrentPupils.Pupil) {
  
    if ($YearGroups -contains $entry.NCYear) {
        #Add pupil to users.csv
        if ($entry.EmailAddress) {
            #Year needs to be 2 digit, and 1 is the lowest year, below that use PR/PK/TK/KG/PS1/PS2/PS3/Other
            if ($entry.NCYear -eq "0") {
                $entryYearGroup = "KG"
            } else {
                $entryYearGroup = $entry.NCYear.PadLeft(2, '0')
            }
            Add-Content -Path "$CSVPath\users.csv", -Value ($entry.schoolId + ",$DFENo," + $entry.EmailAddress + ",student,,,," + $entryYearGroup)
        } else {
            $Errors += [PSCustomObject] @{
                "SchoolID" = $entry.schoolId
                "Surname" = $entry.Surname
                "Forename" = $entry.Forename
                "Year" = $entry.NCYear
                "Error" = "No email address"
            }
        }
    }
}


# Create array of all staff, and array of active teachers
$AllStaff = @()
foreach ($entry in $isamsData.iSAMS.HRManager.CurrentStaff.StaffMember) {
    $AllStaff += [PSCustomObject]@{
        "Id"=$entry.Id
        "EmailAddress"=$entry.SchoolEmailAddress
        }
}
$ActiveTeachers = @()
$ExtraTeacherIDs = @()

if ($ExtraTeachers) {
    foreach ($teacher in $ExtraTeachers) {
        $TeacherData = $AllStaff.Where{$_.EmailAddress -like $teacher}
        if ($TeacherData) {
            $ExtraTeacherIDs += $TeacherData.Id
        } else {
            #If the teacher doesn't exist in the source, create an entry for them (e.g. integration service accounts) ID is a string so can just use their email.
            $AllStaff += [PSCustomObject]@{
                "Id"=$teacher
                "EmailAddress"=$teacher
            }
            $ExtraTeacherIDs += $teacher
        }
    }
    $ActiveTeachers += $ExtraTeacherIDs
}


$SetPupils = @{}
# Loop through sets adding pupils to class
foreach ($entry in $isamsData.iSAMS.TeachingManager.SetLists.SetList) {
    [int]$setID = $entry.SetId.InnerText
    if (!$SetPupils.ContainsKey($setID)) {
        $SetPupils.Add($setId, @())
    }
    $SetPupils[$setId] += $entry.SchoolId.InnerText
}


# Get all sets, loop through sets (applying filter if applicable) and build up Classes hash table to contain class name/students/teachers
$Classes = @{}
foreach ($entry in $isamsData.iSAMS.TeachingManager.Sets.Set) {
    $isFiltered = ($ExcludedSubjects -notcontains $entry.SubjectId.InnerText)
    $isYearFiltered = ($YearGroups -contains $entry.YearId.InnerText)

    if ($isYearFiltered -And $isFiltered) {
        $ClassData = @{"SetID" = $entry.Id; "Subject" = $entry.SubjectId.InnerText; "SetCode" = $entry.SetCode + $Suffix; "Year" = $entry.YearId.InnerText}
        $TeacherList = @()
        foreach ($teacher in $entry.Teachers.Teacher) {
            $TeacherList += $teacher.StaffId
            if ($ActiveTeachers -notcontains $teacher.StaffId) {
                $ActiveTeachers += $teacher.StaffId
            }
        }
        $TeacherList += $ExtraTeacherIDs
        $ClassData.Add("Teacher", $TeacherList)
        $ClassData.Add("Pupils", $SetPupils.Get_Item([int]$entry.Id))
        $Classes.Add($entry.Id, $ClassData)
    }
}



#Work through each set

foreach ($Class in $Classes.Keys) {
    $thisClassData = $Classes[$Class]
    if ($thisClassData.Pupils.Count -gt 0) {
        # Write class to classes.csv
        Add-Content -Path "$CSVPath\classes.csv" -Value ($thisClassData.SetId + "," + $DFENo + "," + $thisClassData.SetCode)
        # Add members
        foreach ($teacher in $thisClassData.Teacher) {
            #write teacher to enrollments
            Add-Content -Path "$CSVPath\enrollments.csv" -Value ($thisClassData.SetId + "," + $teacher + ",teacher")
        }
        foreach ($pupil in $thisClassData.Pupils) {
            #write pupil to enrollments
            Add-Content -Path "$CSVPath\enrollments.csv" -Value ($thisClassData.SetId + "," + $pupil + ",student")
        }
    }
}
    
foreach ($teacher in $ActiveTeachers) {
    $TeacherData = $AllStaff.Where{$_.Id -eq $teacher}
    if ($TeacherData) {
        #Write teachers to users.csv
        Add-Content -Path "$CSVPath\users.csv", -Value ($TeacherData.Id + ",$DFENo," + $TeacherData.EmailAddress + ",teacher,,,,")
    }
}
foreach ($Error in $Errors) {
    Write-Output $Error
}
