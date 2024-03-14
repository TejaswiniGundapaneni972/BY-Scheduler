function Send-Email {
    [CmdletBinding()]
    [OutputType([bool])]
    param (

        [Parameter(Mandatory = $True)]
        [ValidateSet("Phasewise", "FinalReport", "LongRunningActivity", "Failure", "Cancellation", "ReSchedule")]
        [string]$EmailScenario,

        [Parameter(Mandatory = $false)]
        [ValidateSet("PreFailoverValidation", "Failover", "DataSyncToPrimary", "PreFailbackValidation", "Failback", "DataSyncToDR", "PostFailbackValidation")]
        [System.String]$DRPhase,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [guid]$RunId,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$SharepointLocation,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$FileName,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$ErrorMessage,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$TaskId

    )

    begin {

        # inform that function has started
        Write-Information -MessageData "`n`"$($MyInvocation.MyCommand.Name)`" function has started..."

        # print passed parameters
        Write-Information -MessageData "Printing received parameters..."
        Write-Information -MessageData $($($MyInvocation.BoundParameters | Out-String) -replace "`n$")

    }

    process {

        # Read Sharepoint credentials from the Data.json
        Write-Information -MessageData "Getting SMTP Creds and Email template Values from Data.json..."
        try {
            $SMTPUsername = Get-ValueFromJson -ModuleID "Module15" -TaskID "SMTPServerDetails" -KeyName "SMTPUsername"
            $ToEmailAddress = Get-ValueFromJson -ModuleID "Module15" -TaskID "EmailTemplates" -KeyName "ToEmailAddress"
            $FromEmailAddress = Get-ValueFromJson -ModuleID "Module15" -TaskID "EmailTemplates" -KeyName "FromEmailAddress"
            $APIURL = Get-ValueFromJson -ModuleID "Module15" -TaskID "SMTPServerDetails" -KeyName "APIURL"
            $EmailTemplate = Get-ValueFromJson -ModuleID "Module15" -TaskID "EmailTemplates" -KeyName $EmailScenario
        }
        catch {
            Write-Information -MessageData "Encountered Error while getting Data from Json within Module.`n" -InformationAction Continue
            Write-Information -MessageData $($_.Exception | Out-String) -InformationAction Continue; Write-Information -MessageData $($_.InvocationInfo | Out-String) -InformationAction Continue; throw
        }

        # Set Email body and Subject
        if($EmailScenario -eq "Phasewise")
        {
            $EmailBody = $EmailTemplate.body -f $DRPhase
            $EmailSubject = $EmailTemplate.Subject -f $RunId
        }
        elseif ($EmailScenario -eq "FinalReport") {
            $EmailBody = $EmailTemplate.body -f $SharepointLocation, $FileName
            $EmailSubject = $EmailTemplate.Subject -f $RunId
        }
        elseif ($EmailScenario -eq "LongRunningActivity") {
            $EmailBody = $EmailTemplate.body
            $EmailSubject = $EmailTemplate.Subject -f $RunId
        }
        elseif ($EmailScenario -eq "Failure") {
            $EmailBody = $EmailTemplate.body -f $DRPhase, $TaskId, $ErrorMessage
            $EmailSubject = $EmailTemplate.Subject -f $RunId, $DRPhase
        }
        else {
            $EmailBody = $EmailTemplate.body
            $EmailSubject = $EmailTemplate.Subject -f $RunId
        }

        # Format the email body
        $EmailBody = $EmailBody -replace "`n", "\n"

        # Getting Secret Value from Azure Keyvault[SMTP Secret Value]
        Write-Information -MessageData "Getting Secret Value from Azure Keyvault[SMTP Secret Value]"

        $SMTPPwd = "sk_nVDq4jqvtQs#nU+cJjcG-&XYdZup8tDUMumkK"
      

        # Encode the credentials as Base64
        Write-Information -MessageData "Generating the token for authentication"
        $base64AuthInfo = [System.Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $SMTPUsername, $SMTPPwd)))

        # Define the headers with the Basic Auth token
        $headers = @{
            Authorization = "Basic $base64AuthInfo"
            "Content-Type"  = "application/json"
        }

        #Define body for the request
        $body = @"
        {
        "sandboxMode`": false,
        "messages": [
        {
        "from": {
        "email": `"$FromEmailAddress"
        },
        `"to`": [
        {
            "email": `"$ToEmailAddress"
        }
        ],
        "subject": "$EmailSubject",
        "text": `"$EmailBody`"
        }
    ]
    }
"@

        # Make a API call to send email
        Write-Information -MessageData "Triggering the email to the team"

        try {
            $response = Invoke-WebRequest -Uri $APIURL -Method 'POST' -Headers $headers -Body $body -UseBasicParsing -ErrorAction Stop
        }
        catch {
            Write-Information -MessageData "Encountered Error while triggering the email.`n" -InformationAction Continue
            Write-Information -MessageData $($_.Exception | Out-String) -InformationAction Continue; Write-Information -MessageData $($_.InvocationInfo | Out-String) -InformationAction Continue; throw
        }
    }

    end {
        if ($response.StatusCode -eq 201) {
            $TriggeredEmailStatus = $True
        }
        else {
            $TriggeredEmailStatus = $False
        }
        return $TriggeredEmailStatus
    }
}


function Get-ValueFromJson {
    [OutputType([System.Object[]])]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $False)]
        [string]$FilePath = "Data/Data.json",

        [Parameter(Mandatory = $True)]
        [string]$ModuleID,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$TaskID,

        [Parameter(Mandatory = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$KeyName
    )

    begin {
        ## Validating Loging Switch.
        $MessageLogging = $ENV:MESSAGE_LOGGING
        if ($MessageLogging -eq $True) {
            Write-Information -MessageData "`n`"$($MyInvocation.MyCommand.Name)`" function has started..."
            # print passed parameters
            Write-Information -MessageData "Printing received parameters..."
            Write-Information -MessageData $($($MyInvocation.BoundParameters | Out-String) -replace "`n$")
        }
    }

    process {
        ##Fetching out the values using key
        if ($ENV:VERBOSE_LOGGING -eq $True) {
            Write-Information -MessageData "Fetching out the Key from Data.Json"
        }
        try {
            $GetFileData = (Get-Content -path $FilePath -ErrorAction Stop | convertfrom-json)
            $GetValue = $GetFileData.$ModuleID.$TaskID.$KeyName
        }
        catch {
            Write-Information -MessageData "Encountered Error while getting Data from Json.`n" -InformationAction Continue
            Write-Information -MessageData $($_.Exception | Out-String) -InformationAction Continue; Write-Information -MessageData $($_.InvocationInfo | Out-String) -InformationAction Continue; throw
        }
    }

    end {
        if ($GetValue) {
            if ($ENV:VERBOSE_LOGGING -eq $True) {
                Write-Information -MessageData "Value successfully fetched of $KeyName from DataJson"
            }
        }
        else {
            if ($ENV:VERBOSE_LOGGING -eq $True) {
                Write-Information -MessageData "Value not fetched for key: $KeyName from DataJson"
            }
        }
        return $GetValue
    }
}


$EmailScenario ="Phasewise"
$RunId = "366b88da-3a5c-411c-8d39-1e8f78ef5a76"
$DRPhase = "Failover"
Send-Email -RunId $RunId -EmailScenario $EmailScenario -DRPhase $DRPhase
