# Private functions

function Get-OAToken
{
    [CmdletBinding()]

    $token_path = Join-Path -Path $MyInvocation.PSScriptRoot -ChildPath "auth.xml"
    $auth_token = (Import-Clixml $token_path).GetNetworkCredential().Password
    Write-Output $auth_token
}

function Invoke-OAAPIRequest
{
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory,HelpMessage="Which resource are you conecting to? E.g. students, applicants")]
        [string]$Resource,

        [parameter(Mandatory,HelpMessage="Request method is required. E.g. GET, POST, PUT")]
        [ValidateSet("GET","POST","PUT")]
        [string]$Method,

        [parameter()]
        [string]$ContentType="application/json",

        [parameter()]
        [hashtable]$Body
    )

    $BaseURL = "https://cranleigh.openapply.com/api/v1/"
    $URI = $BaseURL + $Resource
    if(-not($Body))
    {
        $Body = @{}
    }
    $Body["auth_token"]="f4e5402b5c7732f7c5fd262522992630"

    $Params=@{
        Uri = $URI
        Method=$Method
        Headers = @{"Content-Type"=$ContentType}
        Body = $Body
    }

    try
    {
        $response = Invoke-RestMethod @Params -ErrorAction Stop
    }
    catch
    {
        Write-Verbose "Hello!"
        Write-Error "$URI ($Method). $($_.Exception.Message)"
        if($Body) {Write-Verbose $Body}
    }

    Write-Output $response

}

function Get-OAStatusCode
{
    [CmdletBinding()]
    param(
        [string]$Status
    )

    $StatusCode = switch($Status)
    {
        pending     {"10"}
        applied     {"20"}
        admitted    {"30"}
        wait-listed {"40"}
        declined    {"50"}
        enrolled    {"60"}
        graduated   {"70"}
        withdrawn   {"80"}
    }

    Write-Output $StatusCode
}

function Get-OAAdmissionsStatus
{
    [CmdletBinding()]
    param(
        [string]$Status,
        [string]$StatusLevel
    )

    $AdmissionsStatus = switch($Status)
    {
        pending     {"Pending"}
        applied     {
            if([string]::IsNullOrEmpty($StatusLevel)){"Applied"}
            else {$StatusLevel}
        }
        wait_listed {"Wait-listed"}
        declined    {"Place Not Offered"}
        withdrawn   {"Withdrawn"}
        admitted    {"Place Accepted"}
        enrolled    {"Current Pupil"}
        graduated   {"Former Pupil"}
    }

    Write-Output $AdmissionsStatus
}

function Get-OAYearGroup
{
    param(
        [ValidateSet('FS1','FS2','Year 1','Year 2','Year 3','Year 4','Year 5','Year 6','Year 7','Year 8','Year 9','Year 10','Year 11','Year 12','Year 13')]
        [string]$Grade
    )
    $YearGroup = switch($Grade)
    {
        "FS1"     {-1}
        "FS2"     {0}
        "Year 1"  {1}
        "Year 2"  {2}
        "Year 3"  {3}
        "Year 4"  {4}
        "Year 5"  {5}
        "Year 6"  {6}
        "Year 7"  {7}
        "Year 8"  {8}
        "Year 9"  {9}
        "Year 10" {10}
        "Year 11" {11}
        "Year 12" {12}
        "Year 13" {13}
    }

    Write-Output $YearGroup
}

function Get-OANationality
{
    [CmdletBinding()]
    param(
        [string]$Nationality
    )

    $Nationality = Switch($Nationality)
    {
        "American (Northern Mariana Islands)" {"American"}
        "American (United States Minor Outlying Islands)" {"American"}
        "American (United States)" {"American"}
        "Chinese (China)" {"Chinese"}
        "Chinese (Hong Kong)" {"Chinese"}
        "Chinese (Macao)" {"Chinese"}
        "Dutch (Bonaire, Sint Eustatius and Saba)" {"Dutch"}
        "Dutch (Netherlands)" {"Dutch"}
        "Dutch (Sint Maarten)" {"Dutch"}
        "French (France)" {"French"}
        "French (French Southern Territories)" {"French"}
        "French (Martinique)" {"French"}
        "French (Mayotte)" {"French"}
        "French (Reunion)" {"French"}
        "French (Saint Pierre And Miquelon)" {"French"}
        "French Guiana" {"French"}
        "French Polynesian" {"French"}
        "Indian (British Indian Ocean Territory)" {"Indian"}
        "Indian (India)" {"Indian"}
        "Italian (Holy See (Vatican City State))" {"Italian"}
        "Italian (Italy)" {"Italian"}
        "Nigerian (Niger)" {"Nigerian"}
        "Nigerian (Nigeria)" {"Nigerian"}
        "Norwegian (Norway)" {"Norwegian"}
        "Norwegian (Svalbard And Jan Mayen)" {"Norwegian"}
        "Swedish (Aland Islands)" {"Swedish"}
        "Swedish (Sweden)" {"Swedish"}
        Default {$Nationality}
    }

    Write-Output $Nationality
}

# Public functions

function Get-OAStudent
{

    [CmdletBinding()]
    param(
        [int[]]$ID,

        [Parameter(HelpMessage="Must be an OpenApply-supported status: Pending | Applied | Admitted | Enrolled | Waitlisted | Declined | Withdrawn")]
        [ValidateSet($null,'Pending','Applied','Admitted','Enrolled','Wait-listed','Declined','Withdrawn')]
        [string]$Status,

        [int]$PerPage=100,

        [boolean]$AllPages=$true
    )

    $Resource = "students"
    $Method = "GET"

    # If $ID exists, return a single record per member of the array:
    if($ID)
    {
        foreach($i in $ID)
        {
            try
            {
                $Data  = Invoke-OAAPIRequest -Resource "$Resource/$i" -Method $Method -ErrorAction Stop
            }
            catch
            {
                Write-Error "Failed to retrieve ID $ID. $($_.Exception.Message)"
            }
            $Student = $Data.student
            Write-Output $Student
        }
    }
    # Otherwise return a set of records, looping through if necessary:
    else
    {
        # Prepare API parameters for splatting:
	    $Params=@{
		    Resource=$Resource
            Method=$Method
		    Body = @{
			    status=$Status
			    count=$PerPage
		    }
	    }

        # The first API request returns $PerPage records, along with meta-data which provides the overall number of pages
        try
        {
            $Data  = Invoke-OAAPIRequest @Params -ErrorAction Stop
            $Pages = $Data.meta.pages
        }
        catch
        {
            Write-Warning "Failed to connect to API. $($_.Exception.Message)"
        }


        # Initialise a blank array and add the first set of student records from $Data.students
        # $Students = @()
        $Students = $Data.students

        Write-Verbose "Page 1 of $Pages. Currently holding $($Students.Count) records"

        #If $AllPages is set to $true, then retrieve all available records from OpenApply
        if($AllPages)
        {
            # Use for-loop and $Params["page"] key to retrieve all records and add them to the $Students array.
            # Page 1 was returned in the first request, so the loop starts from page 2
            for($i=2; $i -le $Pages; $i++){
                $Params['Body']['page']=$i

                Write-Verbose "Requesting page $($Params['Body']['page']) or $Pages"
                try
                {
                    $Data = Invoke-OAAPIRequest @Params -ErrorAction Stop
                }
                catch
                {
                    Write-Warning "Failed to connect to API. $($_.Exception.Message)"
                }
                $Students += $Data.students

                Write-Verbose "Retrieved page $i of $Pages. Currently holding $($Students.Count) student records"
            }
        }

        Write-Verbose "Finished: total result size $($Students.Count) records"
        Write-Output $Students
    }

}

function Set-OAStudent
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int[]]$ID,

        [Parameter(HelpMessage="Must be an OpenApply-supported status: Pending | Applied | Admitted | Enrolled | Waitlisted | Declined | Withdrawn")]
        [ValidateSet($null,'Pending','Applied','Admitted','Enrolled','Wait-listed','Declined','Withdrawn')]
        [string]$Status
    )

    $BaseURL = "https://cranleigh.openapply.com/api/v1/students"
    $Token = (Get-OAToken)
    $StatusCode = Get-OAStatusCode -Status $Status

    Write-Verbose "Status Code = $StatusCode"

    # Prepare API parameters for splatting:
	$Params=@{
		Uri="$BaseURL/$ID/status"
        Method="PUT"
        Headers = @{
            "Content-Type"='application/x-www-form-urlencoded'
        }
        Body = @{
			auth_token=$Token
			status=$StatusCode
		}
	}

    # Process the API request:
    try
    {
        $Response = Invoke-RestMethod @Params -ErrorAction Stop
        Write-Verbose "Updated status: $($Response.student.status)"
    }
    catch
    {
        # Error handling code courtesy of Chris Wahl:
        # http://wahlnetwork.com/2015/02/19/using-try-catch-powershells-invoke-webrequest/
        $error_result = $_.Exception.Response.GetResponseStream()
        $error_reader = New-Object System.IO.StreamReader($error_result)
        $error_response_body = $error_reader.ReadToEnd();
        $error_message = ($error_response_body | ConvertFrom-Json).errors
        Write-Warning "API returned an error: ""$error_message"""
    }
}

function Invoke-OAiSAMSSync
{
    #requires -module iSAMS

    [CmdletBinding()]
    param(
        [string]$LogFile = "C:\Temp\oasync.log",

        [string[]]$Status = @('applied','wait-listed','admitted'),

        [int[]]$ID
    )

    $Warnings=@()
    $students=@()

    if($ID)
    {
        foreach($i in $ID)
        {
            $students += Get-OAStudent -ID $i
        } #foreach
    } #if
    elseif($Status)
    {
        foreach($s in $Status)
        {
            $students += Get-OAStudent -Status $s
        } #foreach
    } #if

    foreach($oas in $students)
    {
        $Name = "$($oas.first_name) $($oas.last_name) ($($oas.id))"

        # We need an iSAMS ID to continue:
        if($oas.student_id)
        {
            if($oas.student_id.length -ne 12)
            {
                $Warnings += "Invalid iSAMS ID for $Name. Moving to next record."
                continue
            } #if

            # If there is no DOB provided, write a warning and move onto the next student:
            if([string]::IsNullOrEmpty($oas.birth_date))
            {
                $Warnings += "No date of birth for $Name. Moving to next record."
                continue
            } #if

            # If preferred name isn't provided, use the first name:
            if([string]::IsNullOrEmpty($oas.preferred_name))
            {
                $oas.preferred_name = $oas.first_name
            } #if

            Write-Verbose "Setting properties for $Name to be applied to iSAMS SchoolID $($oas.student_id)"
            $props = @{
                SchoolId                 = $oas.student_id
                forename                 = $oas.first_name
                middleNames              = $oas.custom_fields.middle_name_s
                surname                  = $oas.last_name
                preferredName            = $oas.preferred_name
                dateOfBirth              = [datetime]$oas.birth_date
                gender                   = switch($oas.gender){ "female" {"F"} "male" {"M"} }
                admissionStatus          = Get-OAAdmissionsStatus -Status $($oas.status) -StatusLevel $($oas.status_level)
                nationalities            = @(Get-OANationality -Nationality $oas.nationality)
                enrolmentSchoolYearGroup = Get-OAYearGroup -Grade $oas.grade
                enrolmentSchoolYear      = $oas.enrollment_year
            } #hashtable
            Set-iSAMSApplicant @props -Verbose:($PSBoundParameters.ContainsKey('Verbose')) -WarningVariable +Warnings

        } #if
        else
        {
            $Warnings += "No iSAMS ID for $Name"
        } #else
    } #foreach
    $Warnings > $LogFile
}
