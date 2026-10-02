<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 26.04.24.2052

<#
.DESCRIPTION
This Exchange Online script runs the Get-CalendarDiagnosticObjects script and returns a summarized timeline of actions in clear English
as well as the Calendar Diagnostic Objects in Excel.

.PARAMETER Identity
One or more SMTP Address of EXO User Mailbox to query.

.PARAMETER Subject
Subject of the meeting to query, only valid if Identity is a single user.

.PARAMETER MeetingID
The MeetingID of the meeting to query.

.PARAMETER TrackingLogs
Include specific tracking logs in the output. Only usable with the MeetingID parameter. Collected by default; use -NoTrackingLogs to skip.

.PARAMETER NoTrackingLogs
Do not collect Tracking Logs.

.PARAMETER Exceptions
Include Exception objects in the output. Collected by default for direct MeetingID searches and for Subject searches that resolve to exactly one MeetingID; use -NoExceptions to skip.

.PARAMETER ExportToExcel
Export the output to an Excel file with formatting.  Running the script for multiple users will create multiple tabs in the Excel file. (Default)

.PARAMETER ExportToCSV
Export the output to 3 CSV files per user.

.PARAMETER CaseNumber
Case Number to include in the Filename of the output.

.PARAMETER ShortLogs
Limit Logs to 500 instead of the default 2000, in case the server has trouble responding with the full logs.

.PARAMETER MaxLogs
Increase log limit to 12,000 in case the default 2000 does not contain the needed information. Note this can be time consuming, and it does not contain all the logs such as User Responses.

.PARAMETER CustomProperty
Advanced users can add custom properties to the output in the RAW output. This is not recommended unless you know what you are doing. The properties must be in the format of "PropertyName1, PropertyName2, PropertyName3". The properties will only be added to the RAW output.

.PARAMETER ExceptionDate
Date of the Exception Meeting to collect logs for. Fastest way to get Exceptions for a meeting after you have identified the MeetingID.

.PARAMETER NoExceptions
Do not collect Exception Meetings.  This was the default behavior of the script, now exceptions are collected by default.

.PARAMETER FastExceptions
Use the AppointmentRecurrenceBlob to find Exception dates, then collect them with -ExceptionDate. Fast exception collection uses the last 6 months by default. If parsing fails, the script falls back to the legacy per-appointment collector.

.PARAMETER AllExceptions
When using -FastExceptions, collect all Exception dates instead of the default last 6 months.

.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity someuser@microsoft.com -MeetingID 040000008200E00074C5B7101A82E008000000008063B5677577D9010000000000000000100000002FCDF04279AF6940A5BFB94F9B9F73CD
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity someuser@microsoft.com -Subject "Test One Meeting Subject"
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity someuser@microsoft.com -Subject "Test One Meeting Subject" -NoExceptions
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity User1, User2, Delegate -MeetingID $MeetingID
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity $Users -MeetingID $MeetingID -NoExceptions
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity $Users -MeetingID $MeetingID -NoTrackingLogs
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity $Users -MeetingID $MeetingID -ExportToExcel -CaseNumber 123456
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity $Users -MeetingID $MeetingID -ExceptionDate "01/28/2024" -CaseNumber 123456
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity $Users -MeetingID $MeetingID -FastExceptions
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity $Users -MeetingID $MeetingID -FastExceptions -AllExceptions

.SYNOPSIS
Used to collect easy to read Calendar Logs.

.LINK
    https://aka.ms/callogformatter
#>

[CmdletBinding(DefaultParameterSetName = 'Subject',
    SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory, Position = 0, HelpMessage = "Enter the Identity of the mailbox(es) to query. Press <Enter> again when done.")]
    [string[]]$Identity,
    [Parameter(HelpMessage = "Export all Logs to Excel (Default).")]
    [switch]$ExportToExcel,
    [Parameter(HelpMessage = "Export all Logs to CSV files.")]
    [switch]$ExportToCSV,
    [Parameter(HelpMessage = "Case Number to include in the Filename of the output.")]
    [string]$CaseNumber,
    [Parameter(HelpMessage = "Limit Logs to 500 instead of the default 2000, in case the server has trouble responding with the full logs.")]
    [switch]$ShortLogs,
    [Parameter(HelpMessage = "Limit Logs to 12000 instead of the default 2000, in case the server has trouble responding with the full logs.")]
    [switch]$MaxLogs,
    [Parameter(HelpMessage = "Custom Property to add to the RAW output.")]
    [string[]]$CustomProperty,

    [Parameter(Mandatory, ParameterSetName = 'MeetingID', Position = 1, HelpMessage = "Enter the MeetingID of the meeting to query. Recommended way to search for CalLogs.")]
    [string]$MeetingID,
    [Parameter(HelpMessage = "Include specific tracking logs in the output. Only usable with the MeetingID parameter.")]
    [switch]$TrackingLogs,
    [Parameter(HelpMessage = "Do Not collect Tracking Logs.")]
    [switch]$NoTrackingLogs,
    [Parameter(HelpMessage = "Include Exception objects in the output. Subject searches also collect them when exactly one MeetingID is found.")]
    [switch]$Exceptions,
    [Parameter(HelpMessage = "Date of the Exception to collect the logs for after identifying the MeetingID.")]
    [DateTime]$ExceptionDate,
    [Parameter(HelpMessage = "Do Not collect Exception Meetings.")]
    [switch]$NoExceptions,
    [Parameter(HelpMessage = "Use AppointmentRecurrenceBlob to collect exceptions by ExceptionDate first, then fall back to the legacy path if parsing fails.")]
    [switch]$FastExceptions,
    [Parameter(HelpMessage = "When using FastExceptions, collect all Exception dates instead of the default last 6 months.")]
    [switch]$AllExceptions,

    [Parameter(Mandatory, ParameterSetName = 'Subject', Position = 1, HelpMessage = "Enter the Subject of the meeting. Do not include the RE:, FW:, etc.,  No wild cards (* or ?)")]
    [string]$Subject
)

# ===================================================================================================
# Auto update script
# ===================================================================================================
$BuildVersion = "26.04.24.2052"




function Confirm-ProxyServer {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $TargetUri
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    try {
        $proxyObject = ([System.Net.WebRequest]::GetSystemWebProxy()).GetProxy($TargetUri)
        if ($TargetUri -ne $proxyObject.OriginalString) {
            Write-Verbose "Proxy server configuration detected"
            Write-Verbose $proxyObject.OriginalString
            return $true
        } else {
            Write-Verbose "No proxy server configuration detected"
            return $false
        }
    } catch {
        Write-Verbose "Unable to check for proxy server configuration"
        return $false
    }
}

function WriteErrorInformationBase {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0],
        [ValidateSet("Write-Host", "Write-Verbose")]
        [string]$Cmdlet
    )

    [string]$errorInformation = [System.Environment]::NewLine + [System.Environment]::NewLine +
    "----------------Error Information----------------" + [System.Environment]::NewLine

    if ($null -ne $CurrentError.OriginInfo) {
        $errorInformation += "Error Origin Info: $($CurrentError.OriginInfo.ToString())$([System.Environment]::NewLine)"
    }

    $errorInformation += "$($CurrentError.CategoryInfo.Activity) : $($CurrentError.ToString())$([System.Environment]::NewLine)"

    if ($null -ne $CurrentError.Exception -and
        $null -ne $CurrentError.Exception.StackTrace) {
        $errorInformation += "Inner Exception: $($CurrentError.Exception.StackTrace)$([System.Environment]::NewLine)"
    } elseif ($null -ne $CurrentError.Exception) {
        $errorInformation += "Inner Exception: $($CurrentError.Exception)$([System.Environment]::NewLine)"
    }

    if ($null -ne $CurrentError.InvocationInfo.PositionMessage) {
        $errorInformation += "Position Message: $($CurrentError.InvocationInfo.PositionMessage)$([System.Environment]::NewLine)"
    }

    if ($null -ne $CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
        $errorInformation += "Remote Position Message: $($CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage)$([System.Environment]::NewLine)"
    }

    if ($null -ne $CurrentError.ScriptStackTrace) {
        $errorInformation += "Script Stack: $($CurrentError.ScriptStackTrace)$([System.Environment]::NewLine)"
    }

    $errorInformation += "-------------------------------------------------$([System.Environment]::NewLine)$([System.Environment]::NewLine)"

    & $Cmdlet $errorInformation
}

function Write-VerboseErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Verbose"
}

function Write-HostErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Host"
}

function Invoke-WebRequestWithProxyDetection {
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Default")]
        [string]
        $Uri,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [switch]
        $UseBasicParsing,

        [Parameter(Mandatory = $true, ParameterSetName = "ParametersObject")]
        [hashtable]
        $ParametersObject,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [string]
        $OutFile
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    if ([System.String]::IsNullOrEmpty($Uri)) {
        $Uri = $ParametersObject.Uri
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Confirm-ProxyServer -TargetUri $Uri) {
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell")
        $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    }

    if ($null -eq $ParametersObject) {
        $params = @{
            Uri     = $Uri
            OutFile = $OutFile
        }

        if ($UseBasicParsing) {
            $params.UseBasicParsing = $true
        }
    } else {
        $params = $ParametersObject
    }

    try {
        Invoke-WebRequest @params
    } catch {
        Write-VerboseErrorInformation
    }
}

<#
    Determines if the script has an update available.
#>
function Get-ScriptUpdateAvailable {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $BuildVersion = "26.04.24.2052"

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $result = [PSCustomObject]@{
        ScriptName     = $scriptName
        CurrentVersion = $BuildVersion
        LatestVersion  = ""
        UpdateFound    = $false
        Error          = $null
    }

    if ((Get-AuthenticodeSignature -FilePath $scriptFullName).Status -eq "NotSigned") {
        Write-Warning "This script appears to be an unsigned test build. Skipping version check."
    } else {
        try {
            $versionData = [Text.Encoding]::UTF8.GetString((Invoke-WebRequestWithProxyDetection -Uri $VersionsUrl -UseBasicParsing).Content) | ConvertFrom-Csv
            $latestVersion = ($versionData | Where-Object { $_.File -eq $scriptName }).Version
            $result.LatestVersion = $latestVersion
            if ($null -ne $latestVersion) {
                $result.UpdateFound = ($latestVersion -ne $BuildVersion)
            } else {
                Write-Warning ("Unable to check for a script update as no script with the same name was found." +
                    "`r`nThis can happen if the script has been renamed. Please check manually if there is a newer version of the script.")
            }

            Write-Verbose "Current version: $($result.CurrentVersion) Latest version: $($result.LatestVersion) Update found: $($result.UpdateFound)"
        } catch {
            Write-Verbose "Unable to check for updates: $($_.Exception)"
            $result.Error = $_
        }
    }

    return $result
}


function Confirm-Signature {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $File
    )

    $IsValid = $false
    $MicrosoftSigningRoot2010 = 'CN=Microsoft Root Certificate Authority 2010, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'
    $MicrosoftSigningRoot2011 = 'CN=Microsoft Root Certificate Authority 2011, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'

    try {
        $sig = Get-AuthenticodeSignature -FilePath $File

        if ($sig.Status -ne 'Valid') {
            Write-Warning "Signature is not trusted by machine as Valid, status: $($sig.Status)."
            throw
        }

        $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
        $chain.ChainPolicy.VerificationFlags = "IgnoreNotTimeValid"

        if (-not $chain.Build($sig.SignerCertificate)) {
            Write-Warning "Signer certificate doesn't chain correctly."
            throw
        }

        if ($chain.ChainElements.Count -le 1) {
            Write-Warning "Certificate Chain shorter than expected."
            throw
        }

        $rootCert = $chain.ChainElements[$chain.ChainElements.Count - 1]

        if ($rootCert.Certificate.Subject -ne $rootCert.Certificate.Issuer) {
            Write-Warning "Top-level certificate in chain is not a root certificate."
            throw
        }

        if ($rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2010 -and $rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2011) {
            Write-Warning "Unexpected root cert. Expected $MicrosoftSigningRoot2010 or $MicrosoftSigningRoot2011, but found $($rootCert.Certificate.Subject)."
            throw
        }

        Write-Host "File signed by $($sig.SignerCertificate.Subject)"

        $IsValid = $true
    } catch {
        $IsValid = $false
    }

    $IsValid
}

<#
.SYNOPSIS
    Overwrites the current running script file with the latest version from the repository.
.NOTES
    This function always overwrites the current file with the latest file, which might be
    the same. Get-ScriptUpdateAvailable should be called first to determine if an update is
    needed.

    In many situations, updates are expected to fail, because the server running the script
    does not have internet access. This function writes out failures as warnings, because we
    expect that Get-ScriptUpdateAvailable was already called and it successfully reached out
    to the internet.
#>
function Invoke-ScriptUpdate {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([boolean])]
    param ()

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $oldName = [IO.Path]::GetFileNameWithoutExtension($scriptName) + ".old"
    $oldFullName = (Join-Path $scriptPath $oldName)
    $tempFullName = (Join-Path ((Get-Item $env:TEMP).FullName) $scriptName)

    if ($PSCmdlet.ShouldProcess("$scriptName", "Update script to latest version")) {
        try {
            Invoke-WebRequestWithProxyDetection -Uri "https://github.com/microsoft/CSS-Exchange/releases/latest/download/$scriptName" -OutFile $tempFullName
        } catch {
            Write-Warning "AutoUpdate: Failed to download update: $($_.Exception.Message)"
            return $false
        }

        try {
            if (Confirm-Signature -File $tempFullName) {
                Write-Host "AutoUpdate: Signature validated."
                if (Test-Path $oldFullName) {
                    Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                }
                Move-Item $scriptFullName $oldFullName
                Move-Item $tempFullName $scriptFullName
                Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                Write-Host "AutoUpdate: Succeeded."
                return $true
            } else {
                Write-Warning "AutoUpdate: Signature could not be verified: $tempFullName."
                Write-Warning "AutoUpdate: Update was not applied."
            }
        } catch {
            Write-Warning "AutoUpdate: Failed to apply update: $($_.Exception.Message)"
        }
    }

    return $false
}

<#
    Determines if the script has an update available. Use the optional
    -AutoUpdate switch to make it update itself. Pass -Confirm:$false
    to update without prompting the user. Pass -Verbose for additional
    diagnostic output.

    Returns $true if an update was downloaded, $false otherwise. The
    result will always be $false if the -AutoUpdate switch is not used.
#>
function Test-ScriptVersion {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Need to pass through ShouldProcess settings to Invoke-ScriptUpdate')]
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $false)]
        [switch]
        $AutoUpdate,
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $updateInfo = Get-ScriptUpdateAvailable $VersionsUrl
    if ($updateInfo.UpdateFound) {
        if ($AutoUpdate) {
            return Invoke-ScriptUpdate
        } else {
            Write-Warning "$($updateInfo.ScriptName) $BuildVersion is outdated. Please download the latest, version $($updateInfo.LatestVersion)."
        }
    }

    return $false
}
if (Test-ScriptVersion -AutoUpdate -VersionsUrl "https://aka.ms/CL-VersionsUrl" -Confirm:$false) {
    # Update was downloaded, so stop here.
    Write-Host -ForegroundColor Red "Script was updated. Please rerun the command."
    return
}

$script:command = $MyInvocation
$script:RunNumber = 0
$script:PreRunErrorCount = $Error.Count
Write-Verbose "The script was started with the following command line:"
Write-Verbose "Name:  $($script:command.MyCommand.name)"
Write-Verbose "Command Line:  $($script:command.line)"
Write-Verbose "Script Version: $BuildVersion"
$script:BuildVersion = $BuildVersion
$script:SubjectSearch = $false
$script:SubjectMeetingIdCount = 0
$script:SubjectResolvedMeetingId = $null
$script:SubjectCanCollectExceptions = $false
$script:SubjectSkippedExceptionCollection = $false
$script:ExceptionCollectionStatus = $null

# ===================================================================================================
# Support scripts
# ===================================================================================================

# ===================================================================================================
# Constants to support the script
# ===================================================================================================
# Non-US date formats for TryParseExact fallback (constant, created once)
[string[]]$script:DateFormats = @(
    "d/M/yyyy h:mm:ss tt",
    "d/M/yyyy H:mm:ss",
    "d/M/yyyy H:mm",
    "d/M/yyyy h:mm tt",
    "dd/MM/yyyy h:mm:ss tt",
    "dd/MM/yyyy H:mm:ss",
    "dd/MM/yyyy HH:mm:ss",
    "dd/MM/yyyy H:mm",
    "dd/MM/yyyy h:mm tt",
    "d/M/yyyy",
    "dd/MM/yyyy"
)

$script:CalendarItemTypes = @{
    'IPM.Schedule.Meeting.Request.AttendeeListReplication' = "AttendeeList"
    'IPM.Schedule.Meeting.Canceled'                        = "Cancellation"
    'IPM.OLE.CLASS.{00061055-0000-0000-C000-000000000046}' = "Exception"
    'IPM.Schedule.Meeting.Notification.Forward'            = "Forward.Notification"
    'IPM.Appointment'                                      = "Ipm.Appointment"
    'IPM.Appointment.Occurrence'                           = "Exception.Occurrence"
    'IPM.Schedule.Meeting.Request'                         = "Meeting.Request"
    'IPM.CalendarSharing.EventUpdate'                      = "SharingCFM"
    'IPM.CalendarSharing.EventDelete'                      = "SharingDelete"
    'IPM.Schedule.Meeting.Resp'                            = "Resp.Any"
    'IPM.Schedule.Meeting.Resp.Neg'                        = "Resp.Neg"
    'IPM.Schedule.Meeting.Resp.Tent'                       = "Resp.Tent"
    'IPM.Schedule.Meeting.Resp.Pos'                        = "Resp.Pos"
    '(Occurrence Deleted)'                                 = "Exception.Deleted"
}

<#
.SYNOPSIS
Resolves the ItemClass to a friendly type name using the CalendarItemTypes table.
If no exact match is found, progressively strips the last dot-segment and retries.
e.g. IPM.Schedule.Meeting.Resp.Tent.Follow → IPM.Schedule.Meeting.Resp.Tent → "Resp.Tent"
#>
function GetItemType {
    param([string]$ItemClass)
    $lookup = $ItemClass
    while (-not [string]::IsNullOrEmpty($lookup)) {
        $type = $script:CalendarItemTypes[$lookup]
        if (-not [string]::IsNullOrEmpty($type)) {
            return $type
        }
        # Strip the last segment and try the parent
        $lastDot = $lookup.LastIndexOf('.')
        if ($lastDot -le 0) { break }
        $lookup = $lookup.Substring(0, $lastDot)
    }
    return $null
}

function IsExceptionItemType {
    param(
        [string]$ItemType
    )

    return (-not [string]::IsNullOrEmpty($ItemType)) -and ($ItemType -like 'Exception*')
}

function IsExceptionLog {
    param(
        $CalLog
    )

    if ($null -eq $CalLog) {
        return $false
    }

    if ($CalLog.IsException -eq $true) {
        return $true
    }

    if ($null -ne $CalLog.CalendarItemType -and ([string]$CalLog.CalendarItemType) -like 'Exception*') {
        return $true
    }

    return IsExceptionItemType (GetItemType $CalLog.ItemClass)
}

function Get-CalendarLogNumericSortValue {
    param(
        $Value
    )

    if ($null -eq $Value) {
        return -1
    }

    $textValue = [string]$Value
    if ([string]::IsNullOrEmpty($textValue) -or $textValue -eq 'NotFound' -or $textValue -eq '-') {
        return -1
    }

    $numericValue = 0
    if ([int]::TryParse($textValue, [ref]$numericValue)) {
        return $numericValue
    }

    return [int]::MaxValue
}

function Get-CalendarLogOriginalStartDateSortRank {
    param(
        $Value
    )

    if ($null -eq $Value) {
        return 0
    }

    $textValue = [string]$Value
    if ([string]::IsNullOrEmpty($textValue) -or $textValue -eq 'NotFound') {
        return 0
    }

    return 1
}

function Get-CalendarLogOriginalStartDateSortValue {
    param(
        $Value
    )

    if ((Get-CalendarLogOriginalStartDateSortRank $Value) -eq 0) {
        return [datetime]::MinValue
    }

    $parsedValue = ConvertDateTimeSilent $Value
    if ($parsedValue -is [datetime]) {
        return $parsedValue
    }

    return [datetime]::MinValue
}

function ConvertValueToString {
    param(
        $Value
    )

    if ($null -eq $Value) {
        return ""
    }

    return [string]$Value
}

function ConvertDateTimeSilent {
    param(
        $Value
    )

    if ($Value -is [datetime]) {
        return $Value
    }

    $textValue = [string]$Value
    if ([string]::IsNullOrEmpty($textValue) -or $textValue -eq 'N/A' -or $textValue -eq 'NotFound') {
        return [datetime]::MinValue
    }

    $InvariantCulture = [System.Globalization.CultureInfo]::InvariantCulture
    $DateStyles = [System.Globalization.DateTimeStyles]::None
    $Parsed = [DateTime]::MinValue

    if ([DateTime]::TryParse($textValue, $InvariantCulture, $DateStyles, [ref]$Parsed)) {
        return $Parsed
    }

    if ([DateTime]::TryParseExact($textValue, $script:DateFormats, $InvariantCulture, $DateStyles, [ref]$Parsed)) {
        return $Parsed
    }

    if ([DateTime]::TryParse($textValue, [ref]$Parsed)) {
        return $Parsed
    }

    return [datetime]::MinValue
}

function Get-CalendarLogTimestampSortValue {
    param(
        $Value
    )

    $parsedValue = ConvertDateTimeSilent $Value
    if ($parsedValue -is [datetime]) {
        return $parsedValue
    }

    return [datetime]::MinValue
}

function Get-CalendarLogDateBucketSortValue {
    param(
        $Value
    )

    $timestamp = Get-CalendarLogTimestampSortValue $Value
    if ($timestamp -eq [datetime]::MinValue) {
        return [datetime]::MinValue
    }

    return $timestamp.Date
}

function Get-CalendarLogTimestampGroupKey {
    param(
        $Value
    )

    $timestamp = Get-CalendarLogTimestampSortValue $Value
    if ($timestamp -eq [datetime]::MinValue) {
        return 'Unknown'
    }

    return $timestamp.ToString('yyyy-MM-dd HH:mm:ss')
}

function Get-SortCheckFailureDetails {
    param(
        [datetime]$PreviousTimestamp,
        [datetime]$CurrentTimestamp,
        [int]$PreviousIndex,
        [int]$CurrentIndex,
        $PreviousEntry,
        $CurrentEntry,
        [string]$ValuePropertyName
    )

    $previousValue = if ($null -ne $PreviousEntry -and -not [string]::IsNullOrEmpty($ValuePropertyName)) { [string]$PreviousEntry.$ValuePropertyName } else { '' }
    $currentValue = if ($null -ne $CurrentEntry -and -not [string]::IsNullOrEmpty($ValuePropertyName)) { [string]$CurrentEntry.$ValuePropertyName } else { '' }
    return "previous[$PreviousIndex]=$PreviousTimestamp raw=[$previousValue]; current[$CurrentIndex]=$CurrentTimestamp raw=[$currentValue]"
}

function Test-LogTimestampOrder {
    param(
        [array]$Entries,
        [string]$Name,
        [string]$ValuePropertyName
    )

    $previousTimestamp = $null
    $previousEntry = $null
    $previousIndex = -1
    $parsedCount = 0
    $skippedCount = 0
    $totalCount = @($Entries).Count

    for ($index = 0; $index -lt @($Entries).Count; $index++) {
        $entry = $Entries[$index]
        if ($null -eq $entry) {
            $skippedCount++
            continue
        }

        $timestamp = Get-CalendarLogTimestampSortValue $entry.$ValuePropertyName
        if ($timestamp -eq [datetime]::MinValue) {
            $skippedCount++
            continue
        }

        $parsedCount++
        if ($null -ne $previousTimestamp -and $previousTimestamp -gt $timestamp) {
            $details = Get-SortCheckFailureDetails -PreviousTimestamp $previousTimestamp -CurrentTimestamp $timestamp -PreviousIndex $previousIndex -CurrentIndex $index -PreviousEntry $previousEntry -CurrentEntry $entry -ValuePropertyName $ValuePropertyName
            Write-Verbose "$Name is not sorted correctly. $details"
            return $false
        }

        $previousTimestamp = $timestamp
        $previousEntry = $entry
        $previousIndex = $index
    }

    $suffix = if ($skippedCount -gt 0) { " Skipped [$skippedCount] row(s) with blank or non-date timestamps." } else { '' }
    Write-Verbose "$Name looks to be sorted correctly. Total rows [$totalCount]. Checked [$parsedCount] timestamped row(s).$suffix"
    return $true
}

function Write-UnknownTimestampDiagnostics {
    param(
        [array]$Entries,
        [string]$Name,
        [string]$ValuePropertyName,
        [int]$MaxItems = 20
    )

    $unknownEntries = [System.Collections.Generic.List[object]]::new()

    for ($index = 0; $index -lt @($Entries).Count; $index++) {
        $entry = $Entries[$index]

        if ($null -eq $entry) {
            [void]$unknownEntries.Add([PSCustomObject]@{
                    Index                = $index
                    Reason               = 'Null entry'
                    RawValue             = ''
                    LogRowType           = ''
                    TriggerAction        = ''
                    ItemClass            = ''
                    CalendarLogRequestId = ''
                    OriginalStartDate    = ''
                })
            continue
        }

        $timestamp = Get-CalendarLogTimestampSortValue $entry.$ValuePropertyName
        if ($timestamp -eq [datetime]::MinValue) {
            [void]$unknownEntries.Add([PSCustomObject]@{
                    Index                = $index
                    Reason               = 'Timestamp resolved to DateTime.MinValue'
                    RawValue             = [string]$entry.$ValuePropertyName
                    LogRowType           = [string]$entry.LogRowType
                    TriggerAction        = [string]$entry.TriggerAction
                    ItemClass            = [string]$entry.ItemClass
                    CalendarLogRequestId = [string]$entry.CalendarLogRequestId
                    OriginalStartDate    = [string]$entry.OriginalStartDate
                })
        }
    }

    if ($unknownEntries.Count -eq 0) {
        Write-Verbose "$Name has no Unknown LogTimestamp rows."
        return
    }

    Write-Verbose "$Name has [$($unknownEntries.Count)] Unknown LogTimestamp row(s) out of [$(@($Entries).Count)] total row(s)."
    foreach ($unknownEntry in ($unknownEntries | Select-Object -First $MaxItems)) {
        Write-Verbose ("  Index [{0}] Reason [{1}] Raw [{2}] LogRowType [{3}] TriggerAction [{4}] ItemClass [{5}] OriginalStartDate [{6}] CalendarLogRequestId [{7}]" -f $unknownEntry.Index, $unknownEntry.Reason, $unknownEntry.RawValue, $unknownEntry.LogRowType, $unknownEntry.TriggerAction, $unknownEntry.ItemClass, $unknownEntry.OriginalStartDate, $unknownEntry.CalendarLogRequestId)
    }

    if ($unknownEntries.Count -gt $MaxItems) {
        Write-Verbose "  ... truncated after [$MaxItems] Unknown rows."
    }
}

function SortCalendarDiagnosticObjectsByTimestamp {
    param(
        [array]$CalLogs
    )

    if ($null -eq $CalLogs -or $CalLogs.Count -le 1) {
        return @($CalLogs)
    }

    return @($CalLogs | Sort-Object @{ Expression = { Get-CalendarLogTimestampSortValue $_.LogTimestamp } })
}

function SortEnhancedCalendarLogsByTimestamp {
    param(
        [array]$CalLogs
    )

    if ($null -eq $CalLogs -or $CalLogs.Count -le 1) {
        return @($CalLogs)
    }

    return @($CalLogs | Sort-Object @{ Expression = { Get-CalendarLogTimestampSortValue $_.LogTimestamp } })
}

function SortCalendarDiagnosticObjects {
    param(
        [array]$CalLogs
    )

    if ($null -eq $CalLogs -or $CalLogs.Count -le 1) {
        return @($CalLogs)
    }

    try {
        return @($CalLogs | Sort-Object
            @{ Expression = { Get-CalendarLogTimestampSortValue $_.LogTimestamp } },
            @{ Expression = { Get-CalendarLogOriginalStartDateSortRank $_.OriginalStartDate } },
            @{ Expression = { Get-CalendarLogOriginalStartDateSortValue $_.OriginalStartDate } })
    } catch {
        Write-Warning "Secondary raw log sort failed; falling back to LogTimestamp-only sorting. $_"
        return SortCalendarDiagnosticObjectsByTimestamp -CalLogs $CalLogs
    }
}

function SortEnhancedCalendarLogs {
    param(
        [array]$CalLogs
    )

    if ($null -eq $CalLogs -or $CalLogs.Count -le 1) {
        return @($CalLogs)
    }

    try {
        # The Enhanced tab is the only output that gets re-ordered.
        # Sort by LogTimestamp first, then use OriginalStartDate only as a tie-breaker for matching timestamps.
        # Rows with unknown LogTimestamp values are preserved and appended to the end.
        $workingLogs = @($CalLogs | Where-Object { $null -ne $_ })
        $removedNullCount = $CalLogs.Count - $workingLogs.Count
        if ($removedNullCount -gt 0) {
            Write-Verbose "Removed [$removedNullCount] null enhanced log row(s) before sorting."
        }

        Write-Verbose "Starting Sorting Date"
        Write-Verbose "Sorting Enhanced logs by LogTimestamp, then OriginalStartDate for matching timestamps."

        $sortIndex = 0
        $sortRows = foreach ($workingLog in $workingLogs) {
            $sortIndex++
            [PSCustomObject]@{
                SortIndex              = $sortIndex
                TimestampValue         = Get-CalendarLogTimestampSortValue $workingLog.LogTimestamp
                OriginalStartDateRank  = Get-CalendarLogOriginalStartDateSortRank $workingLog.OriginalStartDate
                OriginalStartDateValue = Get-CalendarLogOriginalStartDateSortValue $workingLog.OriginalStartDate
                Log                    = $workingLog
            }
        }

        $knownTimestampRows = @($sortRows | Where-Object { $_.TimestampValue -ne [datetime]::MinValue } | Sort-Object TimestampValue, OriginalStartDateRank, OriginalStartDateValue, SortIndex)
        $unknownTimestampRows = @($sortRows | Where-Object { $_.TimestampValue -eq [datetime]::MinValue } | Sort-Object SortIndex)

        if ($unknownTimestampRows.Count -gt 0) {
            Write-Verbose "Keeping [$($unknownTimestampRows.Count)] Unknown LogTimestamp row(s) at the end of the Enhanced output."
        }

        $sortedLogs = @($knownTimestampRows | ForEach-Object { $_.Log }) + @($unknownTimestampRows | ForEach-Object { $_.Log })

        Write-Verbose "Validating Enhanced list order before sub filtering..."
        [void](Test-LogTimestampOrder -Entries $sortedLogs -Name 'Enhanced Tab' -ValuePropertyName 'LogTimestamp')
        Write-UnknownTimestampDiagnostics -Entries $sortedLogs -Name 'Enhanced Tab' -ValuePropertyName 'LogTimestamp'

        return @($sortedLogs)
    } catch {
        Write-Warning "Secondary enhanced log sort failed; falling back to LogTimestamp-only sorting. $_"
        return SortEnhancedCalendarLogsByTimestamp -CalLogs $CalLogs
    }
}

# ===================================================================================================
# Functions to support the script
# ===================================================================================================

<#
.SYNOPSIS
Looks to see if there is a Mapping of ExternalMasterID to FolderName
#>
function MapSharedFolder {
    param(
        $ExternalMasterID
    )
    if ($null -eq $ExternalMasterID -or [string]::IsNullOrEmpty([string]$ExternalMasterID) -or $ExternalMasterID -eq "NotFound") {
        return "Not Shared"
    }

    if ($null -eq $script:SharedFolders -or -not $script:SharedFolders.ContainsKey($ExternalMasterID)) {
        return "UnknownSharedCalendarCopy"
    }

    return $script:SharedFolders[$ExternalMasterID]
}

<#
.SYNOPSIS
Replaces a value of NotFound with a blank string.
#>
function ReplaceNotFound {
    param (
        $Value
    )
    if ($Value -eq "NotFound") {
        return ""
    } else {
        return $Value
    }
}

<#
.SYNOPSIS
Creates a Mapping of ExternalMasterID to FolderName
#>
function CreateExternalMasterIDMap {
    # This function will create a Map of the log folder to ExternalMasterID
    $script:SharedFolders = [System.Collections.SortedList]::new()
    $unknownCount = 0
    Write-Verbose "Starting CreateExternalMasterIDMap"

    foreach ($ExternalID in $script:GCDO.ExternalSharingMasterId | Select-Object -Unique) {
        if ($ExternalID -eq "NotFound") {
            continue
        }

        $AllFolderNames = @($script:GCDO | Where-Object { $_.ExternalSharingMasterId -eq $ExternalID } | Select-Object -ExpandProperty OriginalParentDisplayName | Select-Object -Unique)

        # Default calendar folder names across the top localized versions of Exchange
        # cSpell:ignore Kalender Calendario Calendrier Calendário Календарь
        $DefaultCalendarNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        [void]$DefaultCalendarNames.Add('Calendar')       # English
        [void]$DefaultCalendarNames.Add('Kalender')        # German, Dutch, Swedish, Norwegian, Danish
        [void]$DefaultCalendarNames.Add('Calendario')      # Spanish, Italian
        [void]$DefaultCalendarNames.Add('Calendrier')      # French
        [void]$DefaultCalendarNames.Add('Calendário')     # Portuguese
        [void]$DefaultCalendarNames.Add('Календарь')      # Russian

        # Remove empty/null entries
        $AllFolderNames = @($AllFolderNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

        if ($AllFolderNames.count -gt 1) {
            # We have 2+ FolderNames, remove default calendar folder names (localized) by exact match
            $Filtered = $AllFolderNames | Where-Object { -not $DefaultCalendarNames.Contains($_) }
            if ($Filtered.Count -gt 0) {
                $AllFolderNames = @($Filtered)
            }
            # else keep the original list — all entries were default calendar names
        }

        if ($AllFolderNames.Count -eq 0) {
            $unknownCount++
            $script:SharedFolders[$ExternalID] = "UnknownSharedCalendarCopy$unknownCount"
            Write-Host -ForegroundColor red "Found Zero folder names to map to for ExternalID [$ExternalID]."
        } elseif ($AllFolderNames.Count -eq 1) {
            $script:SharedFolders[$ExternalID] = $AllFolderNames[0]
            Write-Verbose "Found map: [$($AllFolderNames[0])] is for $ExternalID"
        } else {
            # we still have multiple possible Folder Names, need to chose one or combine
            Write-Host -ForegroundColor Red "Unable to Get Exact Folder for $ExternalID"
            Write-Host -ForegroundColor Red "Found $($AllFolderNames.count) possible folders: $($AllFolderNames -join ', ')"

            if ($AllFolderNames.Count -eq 2) {
                $script:SharedFolders[$ExternalID] = $AllFolderNames[0] + " + " + $AllFolderNames[1]
            } else {
                $unknownCount++
                $script:SharedFolders[$ExternalID] = "UnknownSharedCalendarCopy$unknownCount"
            }
        }
    }

    Write-Host -ForegroundColor Green "Created the following Shared Calendar Mapping:"
    foreach ($Key in $script:SharedFolders.Keys) {
        Write-Host -ForegroundColor Green "$Key : $($script:SharedFolders[$Key])"
    }
    Write-Verbose "Created the following Mapping :"
    Write-Verbose $script:SharedFolders
}

<#
.SYNOPSIS
Convert a csv value to multiLine.
#>
function MultiLineFormat {
    param(
        $PassedString
    )
    $PassedString = $PassedString -replace "},", "},`n"
    return $PassedString.Trim()
}

# ===================================================================================================
# Build CSV to output
# ===================================================================================================

<#
.SYNOPSIS
Builds the CSV output from the Calendar Diagnostic Objects
#>
function BuildCSV {

    Write-Host "Starting to Process Calendar Logs..."
    # De-duplicate while preserving the original collection order for RAW output.
    $rawNullCount = @($script:GCDO).Count - @($script:GCDO | Where-Object { $null -ne $_ }).Count
    if ($rawNullCount -gt 0) {
        Write-Host -ForegroundColor Yellow "Removed [$rawNullCount] null raw calendar log row(s) before processing."
    }
    $dedupedCalLogs = New-Object 'System.Collections.Generic.List[object]'
    $seenCalLogKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    $duplicateCount = 0
    foreach ($calLog in @($script:GCDO | Where-Object { $null -ne $_ })) {
        $calLogKey = Get-CalendarDiagnosticObjectDeduplicationKey -CalLog $calLog
        if ($seenCalLogKeys.Add($calLogKey)) {
            [void]$dedupedCalLogs.Add($calLog)
        } else {
            $duplicateCount++
        }
    }
    if ($duplicateCount -gt 0) {
        Write-Host -ForegroundColor Yellow "Removed [$duplicateCount] duplicate Calendar Log entries before processing."
    }
    $script:GCDO = @($dedupedCalLogs)
    $script:MailboxList = @{}
    # Initialize lookup caches to avoid redundant CN resolution across hundreds of log entries
    $script:SMTPAddressCache = @{}
    $script:DisplayNameCache = @{}
    Write-Host "Creating Map of Mailboxes to CNs..."
    CreateExternalMasterIDMap
    ConvertCNtoSMTP
    FixCalendarItemType($script:GCDO)

    Write-Host "Making Calendar Logs more readable..."
    [void](Test-LogTimestampOrder -Entries $script:GCDO -Name 'RAW Tab before sorting' -ValuePropertyName 'LogTimestamp')
    Write-UnknownTimestampDiagnostics -Entries $script:GCDO -Name 'RAW Tab before sorting' -ValuePropertyName 'LogTimestamp'
    $Index = 0
    $GCDOResults = foreach ($CalLog in $script:GCDO) {
        $Index++
        $ItemType = GetItemType $CalLog.ItemClass

        # CleanNotFounds
        $PropsToClean = "FreeBusyStatus", "ClientIntent", "AppointmentSequenceNumber", "AppointmentLastSequenceNumber", "RecurrencePattern", "AppointmentAuxiliaryFlags", "EventEmailReminderTimer", "IsSeriesCancelled", "AppointmentCounterProposal", "MeetingRequestType", "SendMeetingMessagesDiagnostics", "AttendeeCollection"
        foreach ($Prop in $PropsToClean) {
            # Exception objects, etc. don't have these properties.
            if ($null -ne $CalLog.$Prop) {
                $CalLog.$Prop = ReplaceNotFound($CalLog.$Prop)
            }
        }

        # Output one row (collected by the foreach assignment)
        [PSCustomObject]@{
            #'LogRow'                         = $Index
            'LogTimestamp'                   = ConvertDateTime($CalLog.LogTimestamp)
            'LogRowType'                     = ConvertValueToString($CalLog.LogRowType)
            'SubjectProperty'                = $CalLog.SubjectProperty
            'Client'                         = $CalLog.ShortClientInfoString
            'LogClientInfoString'            = $CalLog.LogClientInfoString
            'TriggerAction'                  = $CalLog.CalendarLogTriggerAction
            'ItemClass'                      = $ItemType
            'Seq:Exp:ItemVersion'            = CompressVersionInfo($CalLog)
            'Organizer'                      = GetDisplayName($CalLog.From)
            'From'                           = GetSMTPAddress($CalLog.From)
            'FreeBusy'                       = ConvertValueToString($CalLog.FreeBusyStatus)
            'ResponsibleUser'                = GetSMTPAddress($CalLog.ResponsibleUserName)
            'Sender'                         = GetSMTPAddress($CalLog.Sender)
            'LogFolder'                      = $CalLog.ParentDisplayName
            'OriginalLogFolder'              = $CalLog.OriginalParentDisplayName
            'SharedFolderName'               = MapSharedFolder($CalLog.ExternalSharingMasterId)
            'ReceivedRepresenting'           = GetSMTPAddress($CalLog.ReceivedRepresenting)
            'MeetingRequestType'             = ConvertValueToString($CalLog.MeetingRequestType)
            'StartTime'                      = ConvertDateTime($CalLog.StartTime)
            'EndTime'                        = ConvertDateTime($CalLog.EndTime)
            'OriginalStartDate'              = ConvertDateTime($CalLog.OriginalStartDate)
            'Location'                       = $CalLog.Location
            'CalendarItemType'               = ConvertValueToString($CalLog.CalendarItemType)
            'RecurrencePattern'              = $CalLog.RecurrencePattern
            'AppointmentAuxiliaryFlags'      = ConvertValueToString($CalLog.AppointmentAuxiliaryFlags)
            'DisplayAttendeesAll'            = $(if ($CalLog.DisplayAttendeesAll -eq "NotFound") { "-" } else { $CalLog.DisplayAttendeesAll })
            'AttendeeCount'                  = GetAttendeeCount($CalLog.DisplayAttendeesAll)
            'AppointmentState'               = ConvertValueToString($CalLog.AppointmentState)
            'ResponseType'                   = ConvertValueToString($CalLog.ResponseType)
            'ClientIntent'                   = ConvertValueToString($CalLog.ClientIntent)
            'AppointmentRecurring'           = $CalLog.AppointmentRecurring
            'HasAttachment'                  = $CalLog.HasAttachment
            'IsCancelled'                    = $CalLog.IsCancelled
            'IsAllDayEvent'                  = $CalLog.IsAllDayEvent
            'Sensitivity'                    = $CalLog.Sensitivity
            'IsSeriesCancelled'              = $CalLog.IsSeriesCancelled
            'SendMeetingMessagesDiagnostics' = $CalLog.SendMeetingMessagesDiagnostics
            'AttendeeCollection'             = MultiLineFormat($CalLog.AttendeeCollection)
            'CalendarLogRequestId'           = ConvertValueToString($CalLog.CalendarLogRequestId)    # Move to front.../ Format in groups???
            'CleanGlobalObjectId'            = $CalLog.CleanGlobalObjectId
        }
    }
    [void](Test-LogTimestampOrder -Entries $GCDOResults -Name 'Enhanced Tab before sorting' -ValuePropertyName 'LogTimestamp')
    Write-UnknownTimestampDiagnostics -Entries $GCDOResults -Name 'Enhanced Tab before sorting' -ValuePropertyName 'LogTimestamp'
    # Keep RAW output in collected order; only the Enhanced projection is re-sorted for display and Timeline generation.
    $script:EnhancedCalLogs = SortEnhancedCalendarLogs -CalLogs $GCDOResults
    [void](Test-LogTimestampOrder -Entries $script:GCDO -Name 'RAW Tab' -ValuePropertyName 'LogTimestamp')

    Write-Host -ForegroundColor Green "Calendar Logs have been processed, Exporting logs to file..."
    BuildOrganizerUserNameMap
    Export-CalLog
}

<#
.SYNOPSIS
    Builds a set of all known identifiers (display names and email addresses) for the Organizer.
    This allows change-detection code to treat different representations of the same person as equal.
    Sources: From, Sender, ReceivedBy, ReceivedRepresenting, and AttendeeListDetails from raw CalLogs.
#>
function BuildOrganizerUserNameMap {
    $script:OrganizerIdentities = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    # Seed with the enhanced Organizer display name and From SMTP from the first qualifying log
    $firstOrgLog = $script:EnhancedCalLogs | Where-Object {
        ($_.ItemClass -eq 'Ipm.Appointment' -or $_.ItemClass -like 'Exception*') -and
        $_.SharedFolderName -eq 'Not Shared'
    } | Select-Object -First 1

    if ($null -ne $firstOrgLog) {
        if (-not [string]::IsNullOrEmpty($firstOrgLog.From) -and $firstOrgLog.From -ne '-') {
            [void]$script:OrganizerIdentities.Add($firstOrgLog.From.Trim())
        }
        if (-not [string]::IsNullOrEmpty($firstOrgLog.Organizer) -and $firstOrgLog.Organizer -ne '-') {
            [void]$script:OrganizerIdentities.Add($firstOrgLog.Organizer.Trim())
        }
    }

    # Walk raw CalLogs for additional representations on qualifying rows
    foreach ($CalLog in $script:GCDO) {
        if ($null -eq $CalLog.ItemClass -or
            ((GetItemType $CalLog.ItemClass) -ne 'Ipm.Appointment' -and -not (IsExceptionItemType (GetItemType $CalLog.ItemClass)))) {
            continue
        }

        # From field — extract display name and email
        AddIdentitiesFromCNField $CalLog.From

        # ReceivedBy and ReceivedRepresenting — "Display Name" <email>
        AddIdentitiesFromCNField $CalLog.ReceivedBy
        AddIdentitiesFromCNField $CalLog.ReceivedRepresenting

        # AttendeeListDetails — JSON mapping; look for the organizer's entry by matching known identities
        if ($null -ne $CalLog.AttendeeListDetails -and -not [string]::IsNullOrEmpty([string]$CalLog.AttendeeListDetails) -and [string]$CalLog.AttendeeListDetails -ne 'NotFound') {
            try {
                $attendeeMap = $CalLog.AttendeeListDetails | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($null -ne $attendeeMap) {
                    foreach ($prop in $attendeeMap.PSObject.Properties) {
                        $email = $prop.Name
                        $displayName = $prop.Value.DisplayName
                        # Only add if we already know this email belongs to the organizer
                        if ($script:OrganizerIdentities.Contains($email)) {
                            if (-not [string]::IsNullOrEmpty($displayName)) {
                                [void]$script:OrganizerIdentities.Add($displayName.Trim())
                            }
                        }
                    }
                }
            } catch {
                Write-Verbose "BuildOrganizerUserNameMap: Unable to parse AttendeeListDetails: $_"
            }
        }
    }

    if ($script:OrganizerIdentities.Count -gt 0) {
        Write-Verbose "OrganizerIdentities map contains $($script:OrganizerIdentities.Count) entries:"
        foreach ($id in $script:OrganizerIdentities) { Write-Verbose "`t$id" }
    } else {
        Write-Verbose "OrganizerIdentities map is empty — organizer change detection may be less accurate."
    }
}

<#
.SYNOPSIS
    Extracts display name and email address from a CN-format field value and adds them to the organizer identity set.
#>
function AddIdentitiesFromCNField {
    param($FieldValue)

    if ($null -eq $FieldValue -or [string]::IsNullOrEmpty([string]$FieldValue) -or [string]$FieldValue -eq 'NotFound') {
        return
    }

    $text = [string]$FieldValue

    # Extract SMTP address from <email@domain.com> format
    $smtpMatch = [regex]::Match($text, '<([^>]+@[^>]+)>')
    if ($smtpMatch.Success) {
        [void]$script:OrganizerIdentities.Add($smtpMatch.Groups[1].Value.Trim())
    }

    # Extract display name — text before < or quoted text
    if ($text -match '<') {
        $displayPart = ($text -split '<')[0].Trim().Trim('"').Trim()
        if (-not [string]::IsNullOrEmpty($displayPart)) {
            [void]$script:OrganizerIdentities.Add($displayPart)
        }
    }

    # Also resolve via existing helpers — the resolved SMTP and display name
    $resolvedSmtp = GetSMTPAddress $FieldValue
    if (-not [string]::IsNullOrEmpty($resolvedSmtp) -and $resolvedSmtp -ne '-' -and $resolvedSmtp -ne 'NotFound') {
        [void]$script:OrganizerIdentities.Add($resolvedSmtp.Trim())
    }
    $resolvedDisplay = GetDisplayName $FieldValue
    if (-not [string]::IsNullOrEmpty($resolvedDisplay) -and $resolvedDisplay -ne '-' -and $resolvedDisplay -ne 'NotFound') {
        [void]$script:OrganizerIdentities.Add($resolvedDisplay.Trim())
    }
}

<#
.SYNOPSIS
    Returns $true if two organizer identity values refer to the same person,
    using the OrganizerIdentities map for case-insensitive lookup.
#>
function IsSameOrganizerIdentity {
    param(
        [string]$Value1,
        [string]$Value2
    )

    if ([string]::Equals($Value1, $Value2, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    if ($null -ne $script:OrganizerIdentities -and $script:OrganizerIdentities.Count -gt 0) {
        $v1Known = (-not [string]::IsNullOrEmpty($Value1)) -and $script:OrganizerIdentities.Contains($Value1.Trim())
        $v2Known = (-not [string]::IsNullOrEmpty($Value2)) -and $script:OrganizerIdentities.Contains($Value2.Trim())
        if ($v1Known -and $v2Known) {
            return $true
        }
    }

    return $false
}

function ConvertDateTime {
    param(
        [string] $DateTime
    )
    if ([string]::IsNullOrEmpty($DateTime) -or
        $DateTime -eq "N/A" -or
        $DateTime -eq "NotFound") {
        return ""
    }

    $InvariantCulture = [System.Globalization.CultureInfo]::InvariantCulture
    $DateStyles = [System.Globalization.DateTimeStyles]::None
    $Parsed = [DateTime]::MinValue

    # Priority 1 & 2: InvariantCulture TryParse
    # Handles ISO 8601 and M/d/yyyy (the default EXO output format)
    if ([DateTime]::TryParse($DateTime, $InvariantCulture, $DateStyles, [ref]$Parsed)) {
        return $Parsed
    }

    # Priority 3: Explicit non-US formats via TryParseExact
    # Reached when day > 12 makes InvariantCulture fail (e.g., "22/07/2024")
    if ([DateTime]::TryParseExact($DateTime, $script:DateFormats, $InvariantCulture, $DateStyles, [ref]$Parsed)) {
        return $Parsed
    }

    # Priority 4: Current culture (last resort for unexpected formats)
    if ([DateTime]::TryParse($DateTime, [ref]$Parsed)) {
        return $Parsed
    }

    # Priority 5: Return MinValue so sorting is not broken by mixed types
    Write-Warning "Unable to parse date: [$DateTime]"
    return [DateTime]::MinValue
}

function GetAttendeeCount {
    param(
        [string] $AttendeesAll
    )
    if ($AttendeesAll -ne "NotFound") {
        return ($AttendeesAll -split ';').Count
    } else {
        return "-"
    }
}

<#
.SYNOPSIS
Corrects the CalenderItemType column
#>
function FixCalendarItemType {
    param(
        $CalLogs
    )
    foreach ($CalLog in $CalLogs) {
        if ($CalLog.OriginalStartDate -ne "NotFound" -and ![string]::IsNullOrEmpty($CalLog.OriginalStartDate)) {
            $CalLog.CalendarItemType = "Exception"
            $CalLog.isException = $true
        }
    }
}

function CompressVersionInfo {
    param(
        $CalLog
    )
    [string] $CompressedString = ""
    if ($CalLog.AppointmentSequenceNumber -eq "NotFound" -or [string]::IsNullOrEmpty($CalLog.AppointmentSequenceNumber)) {
        $CompressedString = "-:"
    } else {
        $CompressedString = $CalLog.AppointmentSequenceNumber.ToString() + ":"
    }
    if ($CalLog.AppointmentLastSequenceNumber -eq "NotFound" -or [string]::IsNullOrEmpty($CalLog.AppointmentLastSequenceNumber)) {
        $CompressedString += "-:"
    } else {
        $CompressedString += $CalLog.AppointmentLastSequenceNumber.ToString() + ":"
    }
    if ($CalLog.ItemVersion -eq "NotFound" -or [string]::IsNullOrEmpty($CalLog.ItemVersion)) {
        $CompressedString += "-"
    } else {
        $CompressedString += $CalLog.ItemVersion.ToString()
    }

    return $CompressedString
}

# ===================================================================================================
# BuildTimeline
# ===================================================================================================

<#
.SYNOPSIS
    Tries to builds a timeline of the history of the meeting based on the diagnostic objects.

.DESCRIPTION
    By using the Enhanced diagnostic objects, which are sorted before the Timeline is built, we try to give a high level
    overview of what happened to the meeting. This can be use to get a quick overview of the meeting and
    then you can look into the CalLog in Excel to get more details.

    The timeline will skip a lot of the noise (isIgnorable) in the CalLogs. It skips EBA (Event Based Assistants),
    and other EXO internal processes, which are (99% of the time) not interesting to the end user and just setting
    hidden internal properties (i.e. things like HasBeenIndex, etc.)

    It also skips items from Shared Calendars, which are calendars that have a Modern Sharing relationship setup,
    which creates a replicated copy of another users. If you want to look at the actions this user took on
    another users calendar, you can look at that users Calendar Logs.

.NOTES
    The timeline will never be perfect, but if you see a way to make it more understandable, readable, etc.,
    please let me know or fix it yourself on GitHub.
    I use a iterative approach to building this, so it will get better over time.
#>

function FindOrganizer {
    param (
        $CalLog
    )
    $Script:Organizer = "Unknown"
    if ($null -ne $CalLog.From) {
        if ($null -ne $CalLog.From.SmtpEmailAddress) {
            $Script:Organizer = $($CalLog.From.SmtpEmailAddress)
        } elseif ($null -ne $CalLog.From.DisplayName) {
            $Script:Organizer = $($CalLog.From.DisplayName)
        } elseif ($calLog.From -match "^\s*<") {
            $Script:Organizer = $($CalLog.From -split "<")[1] -replace ">", ""
        } else {
            $Script:Organizer = $($CalLog.From)
        }
    }
    Write-Host -ForegroundColor Green "Setting Organizer to : [$Script:Organizer]"
}

function FindFirstMeeting {
    [array]$IpmAppointments = $script:GCDO | Where-Object { $_.ItemClass -eq "IPM.Appointment" -and $_.ExternalSharingMasterId -eq "NotFound" }
    if ($IpmAppointments.count -eq 0) {
        Write-Host "All CalLogs are from Shared Calendar, getting values from first IPM.Appointment."
        $IpmAppointments = $script:GCDO | Where-Object { $_.ItemClass -eq "IPM.Appointment" }
    }
    if ($IpmAppointments.count -eq 0) {
        if ($script:IsOrganizer) {
            Write-Host -ForegroundColor Red "Warning: Cannot find any IPM.Appointments for the Organizer. Check for the Outlook Bifurcation issue."
        } else {
            Write-Host -ForegroundColor Yellow "Warning: No IPM.Appointments found. This user is not the Organizer, so this may be expected (e.g. delegate without the meeting on their own calendar)."
        }
        Write-Host -ForegroundColor Red "Warning: No IPM.Appointment found. CalLogs start to expire after 31 days."
        return $null
    } else {
        return $IpmAppointments[0]
    }
}

function BuildTimeline {
    $script:TimeLineOutput = [System.Collections.Generic.List[object]]::new()

    $script:FirstLog = FindFirstMeeting
    FindOrganizer($script:FirstLog)

    # Ignorable and items from Shared Calendars are not included in the TimeLine.
    [array]$InterestingCalLogs = $script:EnhancedCalLogs |
        Where-Object { $_.LogRowType -eq "Interesting" -and $_.SharedFolderName -eq "Not Shared" }

    if ($InterestingCalLogs.count -eq 0) {
        Write-Host "All CalLogs are Ignorable, nothing to create a timeline with, displaying initial values."
    } else {
        Write-Host "Found $($script:EnhancedCalLogs.count) Log entries, only the $($InterestingCalLogs.count) Non-Ignorable entries will be analyzed in the TimeLine. `n"
    }

    if ($script:CalLogsDisabled) {
        Write-Host -ForegroundColor Red "Warning: CalLogs are disabled for this user, Timeline / CalLogs will be incomplete."
        return
    }

    Write-DashLineBoxColor "  TimeLine for: [$Identity]",
    "CollectionDate: $($(Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))",
    "ScriptVersion: $Script:BuildVersion",
    "  Subject: $($script:GCDO[0].NormalizedSubject)",
    "  Organizer: $Script:Organizer",
    "  MeetingID: $($script:GCDO[0].CleanGlobalObjectId)"
    [array]$Header = "MeetingID: "+ ($script:GCDO[0].CleanGlobalObjectId)

    CreateMeetingSummary -Time "Calendar Timeline for Meeting" -MeetingChanges $Header
    if ($null -ne $FirstLog) {
        CreateMeetingSummary -Time "Initial Message Values" -Entry $script:FirstLog -LongVersion
    }

    # Look at each CalLog and build the Timeline
    foreach ($CalLog in $InterestingCalLogs) {
        [bool] $script:MeetingSummaryNeeded = $False
        [bool] $script:AddChangedProperties = $False

        $MeetingChanges = CreateTimelineRow
        # Create the Timeline by adding to Time to the generated MeetingChanges
        $Time = "$(ConvertDateTime($CalLog.LogTimestamp))"

        if ($MeetingChanges) {
            if ($script:MeetingSummaryNeeded) {
                CreateMeetingSummary -Time $Time -MeetingChanges $MeetingChanges
                CreateMeetingSummary -Time " " -ShortVersion -Entry $CalLog
            } else {
                CreateMeetingSummary -Time $Time -MeetingChanges $MeetingChanges
                if ($script:AddChangedProperties) {
                    FindChangedProperties
                }
            }
        }

        # Setup Previous log (if current logs is an IPM.Appointment)
        if ($null -ne $CalLog.ItemClass -and
            ((GetItemType $CalLog.ItemClass) -eq "Ipm.Appointment" -or (IsExceptionItemType (GetItemType $CalLog.ItemClass)))) {
            $script:PreviousCalLog = $CalLog
        }
    }

    [void](Test-LogTimestampOrder -Entries $script:TimeLineOutput -Name 'Timeline' -ValuePropertyName 'Time')

    Export-Timeline
}
function Convert-Data {
    param(
        [Parameter(Mandatory = $True)]
        [string[]] $ArrayNames,
        [switch ] $NoWarnings = $False
    )
    $ValidArrays = @()
    $ItemCounts = @()
    $VariableLookup = @{}
    foreach ($Array in $ArrayNames) {
        try {
            $VariableData = Get-Variable -Name $Array -ErrorAction Stop
            $VariableLookup[$Array] = $VariableData.Value
            $ValidArrays += $Array
            $ItemCounts += ($VariableData.Value | Measure-Object).Count
        } catch {
            if (!$NoWarnings) {
                Write-Warning -Message "No variable found for [$Array]"
            }
        }
    }
    $MaxItemCount = ($ItemCounts | Measure-Object -Maximum).Maximum
    $FinalArray = for ($Inc = 0; $Inc -lt $MaxItemCount; $Inc++) {
        $FinalObj = New-Object PsObject
        foreach ($Item in $ValidArrays) {
            $FinalObj | Add-Member -MemberType NoteProperty -Name $Item -Value $VariableLookup[$Item][$Inc]
        }
        $FinalObj
    }

    return $FinalArray
}

# ===================================================================================================
# Write Out one line of the Meeting Summary (Time + Meeting Changes)
# ===================================================================================================
function CreateMeetingSummary {
    param(
        [array] $Time,
        [array] $MeetingChanges,
        $Entry,
        [switch] $LongVersion,
        [switch] $ShortVersion
    )

    $InitialSubject = "Subject: " + $Entry.NormalizedSubject
    $InitialOrganizer = "Organizer: " + $Entry.SentRepresentingDisplayName
    $InitialSender = "Sender: " + $Entry.SentRepresentingDisplayName
    $InitialToList = "To List: " + $Entry.DisplayAttendeesAll
    $InitialLocation = "Location: " + $Entry.Location

    if ($ShortVersion -or $LongVersion) {
        $InitialStartTime = "StartTime: " + $Entry.StartTime.ToString()
        $InitialEndTime = "EndTime: " + $Entry.EndTime.ToString()
    }

    if ($longVersion -and ($Entry.Timezone -ne "")) {
        $InitialTimeZone = "Time Zone: " + $Entry.Timezone
    } else {
        $InitialTimeZone = "Time Zone: Not Populated"
    }

    if ($Entry.AppointmentRecurring) {
        $InitialRecurring = "Recurring: Yes - Recurring"
    } else {
        $InitialRecurring = "Recurring: No - Single instance"
    }

    if ($longVersion -and $Entry.AppointmentRecurring) {
        $InitialRecurrencePattern = "RecurrencePattern: " + $Entry.RecurrencePattern
        $InitialSeriesStartTime = "Series StartTime: " + $Entry.StartTime.ToString() + "Z"
        $InitialSeriesEndTime = "Series EndTime: " + $Entry.StartTime.ToString() + "Z"
        if (!$Entry.ViewEndTime) {
            $InitialEndDate = "Meeting Series does not have an End Date."
        }
    }

    if (!$Time) {
        $Time = $Entry.LogTimestamp
    }

    if (!$MeetingChanges) {
        $MeetingChanges = @()
        $MeetingChanges += $InitialSubject, $InitialOrganizer, $InitialSender, $InitialToList, $InitialLocation, $InitialStartTime, $InitialEndTime, $InitialTimeZone, $InitialRecurring, $InitialRecurrencePattern, $InitialSeriesStartTime , $InitialSeriesEndTime , $InitialEndDate
    }

    if ($ShortVersion) {
        $MeetingChanges = @()
        $MeetingChanges += $InitialToList, $InitialLocation, $InitialStartTime, $InitialEndTime, $InitialRecurring
    }

    $script:TimeLineOutput.AddRange(@(Convert-Data -ArrayNames "Time", "MeetingChanges"))
}

$WellKnownCN_CA = "MICROSOFT SYSTEM ATTENDANT"
$CalAttendant = "Calendar Assistant"
$WellKnownCN_Trans = "MicrosoftExchange"
$Transport = "Transport Service"
<#
.SYNOPSIS
Get the Mailbox for the Passed in Identity.
Might want to extend to do 'Get-MailUser' as well.
.PARAMETER CN of the Mailbox
    The mailbox for which to retrieve properties.
.PARAMETER Organization
    [Optional] Organization to search for the mailbox in.
#>
function GetMailbox {
    param(
        [string]$Identity,
        [string]$Organization,
        [bool]$UseGetMailbox
    )

    if ($Identity -like "*<*") {
        $Identity = $Identity -split "<" | Select-Object -Last 1
        $Identity = $Identity -split ">" | Select-Object -First 1
        Write-Verbose "Refined Identity to [$Identity]"
    }

    $params = @{Identity = $Identity
        ErrorAction      = "SilentlyContinue"
    }

    if ($UseGetMailbox) {
        $Cmdlet = "Get-Mailbox"
        $params.Add("IncludeInactiveMailbox", $true)
    } else {
        $Cmdlet = "Get-Recipient"
    }

    try {
        Write-Verbose "Searching $Cmdlet $(if (-not ([string]::IsNullOrEmpty($Organization))) {"with Org: $Organization"}) for $Identity."

        if (-not ([string]::IsNullOrEmpty($Organization)) -and $script:MSSupport) {
            Write-Verbose "Using Organization parameter"
            $params.Add("Organization", $Organization)
        } elseif (-not ([string]::IsNullOrEmpty($Organization))) {
            Write-Verbose "Using -OrganizationalUnit parameter with $Organization."
            $params.Add("Organization", $Organization)
        }

        Write-Verbose "Running $Cmdlet with params: $($params.Values)"
        $RecipientOutput = & $Cmdlet @params
        Write-Verbose "RecipientOutput: $RecipientOutput"

        if (!$RecipientOutput) {
            Write-Host "Unable to find [$Identity]$(if ($Organization -ne `"`" ) {" in Organization:[$Organization]"})."
            Write-Host "Trying to find a Group Mailbox for [$Identity]..."
            $RecipientOutput = Get-Mailbox -Identity $Identity -ErrorAction SilentlyContinue -GroupMailbox
            if (!$RecipientOutput) {
                Write-Host "Unable to find a Group Mailbox for [$Identity] either."
                return $null
            } else {
                Write-Verbose "Found GroupMailbox [$($RecipientOutput.DisplayName)]"
            }
        }

        if ($null -eq $script:PIIAccess) {
            [bool]$script:PIIAccess = CheckForPIIAccess($RecipientOutput.DisplayName)
        }

        if ($script:PIIAccess) {
            Write-Verbose "Found [$($RecipientOutput.DisplayName)]"
        } else {
            Write-Verbose "No PII Access for [$Identity]"
        }

        return $RecipientOutput
    } catch {
        Write-Error "An error occurred while running ${Cmdlet}: [$_]"
    }
}

<#
.SYNOPSIS
Checks the identities are EXO Mailboxes.
#>
function CheckIdentities {
    if (Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue) {
        Write-Host "Validated connection to Exchange Online..."
    } else {
        Write-Error "Get-Mailbox cmdlet not found. Please validate that you are running this script from an Exchange Management Shell and try again."
        Write-Host "Look at Import-Module ExchangeOnlineManagement and Connect-ExchangeOnline."
        exit
    }

    # See if it is a Customer Tenant running the cmdlet. (They will not have access to Organization parameter)
    $script:MSSupport = [Bool](Get-Help Get-Mailbox -Parameter Organization -ErrorAction SilentlyContinue)
    Write-Verbose "MSSupport: $script:MSSupport"

    Write-Host "Checking for at least one valid mailbox..."
    $IdentityList = @()

    Write-Host "Preparing to check $($Identity.count) Mailbox(es)..."

    foreach ($Id in $Identity) {
        [bool]$script:CalLogsDisabled = $false
        $Account = GetMailbox -Identity $Id -UseGetMailbox $true
        if ($null -eq $Account) {
            # -or $script:MB.GetType().FullName -ne "Microsoft.Exchange.Data.Directory.Management.Mailbox") {
            Write-DashLineBoxColor "`n Error: Mailbox [$Id] not found on Exchange Online.  Please validate the mailbox name and try again.`n" -Color Red
            continue
        }
        if (-not (CheckForPIIAccess($Account.DisplayName))) {
            Write-Host -ForegroundColor DarkRed "No PII access for Mailbox [$Id]. Falling back to SMTP Address."
            $IdentityList += $ID
            if ($null -eq $script:MB) {
                $script:MB = $Account
            }
        } else {
            Write-Host "Mailbox [$Id] found as : $($Account.DisplayName)"
            $IdentityList += $Account.PrimarySmtpAddress.ToString()
            if ($null -eq $script:MB) {
                $script:MB = $Account
            }
        }
        # Need to reset between users!!!
        if ($Account.CalendarVersionStoreDisabled -eq $true) {
            $script:CalLogsDisabled = $true
            Write-Host -ForegroundColor DarkRed "Mailbox [$Id] has CalendarVersionStoreDisabled set to True.  This mailbox will not have Calendar Logs."
            Write-Host -ForegroundColor DarkRed "Some logs will be available for Mailbox [$Id] but they will not be complete."
        }
        if ($Account.RecipientTypeDetails -eq "RoomMailbox" -or $Account.RecipientTypeDetails -eq "EquipmentMailbox") {
            if ($script:PIIAccess -eq $true) {
                $script:Rooms += $Account.PrimarySmtpAddress.ToString()
            } else {
                $script:Rooms += $Id
            }
            Write-Host -ForegroundColor Green "[$Id] is a Room / Equipment Mailbox."
        }
    }

    Write-Verbose "IdentityList: $IdentityList"

    if ($IdentityList.count -eq 0) {
        Write-DashLineBoxColor "`n No valid mailboxes found.  Please validate the mailbox name and try again. `n" Red
        exit
    }

    return $IdentityList
}

<#
.SYNOPSIS
Creates a list of CN that are used in the Calendar Logs, Looks up the Mailboxes and stores them in the MailboxList.
#>
function ConvertCNtoSMTP {
    # Creates a list of CN's that we will do MB look up on
    $CNEntries = @()
    $CNEntries += ($script:GCDO.ResponsibleUserName.ToUpper() | Select-Object -Unique)
    $CNEntries += ($script:GCDO.Sender.ToUpper() | Select-Object -Unique)
    # $CNEntries += ($script:GCDO.SenderEmailAddress.ToUpper() | Select-Object -Unique)
    $CNEntries = $CNEntries | Select-Object -Unique
    Write-Verbose "`t Have $($CNEntries.count) CN Entries to look for..."
    Write-Verbose "CNEntries: "; foreach ($CN in $CNEntries) { Write-Verbose `t`t$CN }

    $Org = $script:MB.OrganizationalUnit.split('/')[-1]

    # Creates a Dictionary of MB's that we will use to look up the CN's
    Write-Verbose "Converting CN entries into SMTP Addresses..."

    foreach ($CNEntry in $CNEntries) {
        Write-Verbose  "Processing CNEntry: [$CNEntry]"
        if ($CNEntry -like $WellKnownCN_CA) {
            $script:MailboxList[$CNEntry] = $CalAttendant
        } elseif ($CNEntry -like $WellKnownCN_Trans) {
            $script:MailboxList[$CNEntry] = $Transport
        } elseif ($CNEntry -match 'cn=([\w,\s.@-]*[^/])$') {
            Write-Verbose  "Using Get MB to look up [$CNEntry]"
            $script:MailboxList[$CNEntry] = (GetMailbox -Identity $CNEntry -Organization $Org)
        } elseif ($CNEntry -like "*<*@*>") {
            Write-Verbose "Passing in SMTP address [$CNEntry] to GetSMTPAddress"
            $script:MailboxList[$CNEntry] = GetSMTPAddress -PassedCN $Cn
        } elseif ($CNEntry -like '*</O=*>') {
            $Cn = $CNEntry -split "<" | Select-Object -Last 1
            $Cn = $Cn -split ">" | Select-Object -First 1
            Write-Verbose  "`t Working on [$CN]"
            $script:MailboxList[$CNEntry] = GetMailboxProp -PassedCN $Cn -Prop "PrimarySmtpAddress"
        } else {
            Write-Verbose "ConvertCNtoSMTP: Passed in Value does not look like a CN or SMTP Address: [$CNEntry]"
            $script:MailboxList[$CNEntry] = $CNEntry
        }
    }
    Write-Verbose "MailboxList: $($script:MailboxList.Count) entries found."

    foreach ($key in $script:MailboxList.Keys) {
        $value = $script:MailboxList[$key]
        Write-Verbose "$key :: $($value.DisplayName)"
    }
}

<#
.SYNOPSIS
Gets DisplayName from a passed in CN that matches an entry in the MailboxList
#>
function GetDisplayName {
    param(
        $PassedCN
    )

    # Use string key for cache lookup
    $cacheKey = [string]$PassedCN
    if ($script:DisplayNameCache.ContainsKey($cacheKey)) {
        return $script:DisplayNameCache[$cacheKey]
    }

    Write-Verbose "GetDisplayName:: Working on [$PassedCN]"
    if ($PassedCN.Properties.Name -contains 'DisplayName') {
        $result = GetMailboxProp -PassedCN $PassedCN -Prop "DisplayName"
    } elseif ($PassedCN -match '<' ) {
        $result = $PassedCN.ToString().split("<")[0].replace('"', '')
    } elseif ($PassedCN -match '\[OneOff') {
        $result = $PassedCN.ToString().split('"')[1]
    } else {
        Write-Verbose "Unable to get the DisplayName for [$PassedCN]"
        $result = $PassedCN
    }

    $script:DisplayNameCache[$cacheKey] = $result
    return $result
}

<#
.SYNOPSIS
Gets SMTP Address from a passed in CN that matches an entry in the MailboxList
#>
function GetSMTPAddress {
    param(
        $PassedCN
    )

    # Use string key for cache lookup
    $cacheKey = [string]$PassedCN
    if ($script:SMTPAddressCache.ContainsKey($cacheKey)) {
        return $script:SMTPAddressCache[$cacheKey]
    }

    #Write-Verbose "GetSMTPAddress:: Working on [$PassedCN]"
    if ($PassedCN -match $WellKnownCN_Trans) {
        $result = $Transport
    } elseif ($PassedCN -match $WellKnownCN_CA) {
        $result = $CalAttendant
    } elseif ($PassedCN -like "*<*@*>") {
        # This is a new format that we are seeing in the Calendar Logs.
        # Example: '"Jon Doe" <Jon.Doe@Contoso.com>'
        $smtpMatch = [regex]::Match($PassedCN, '<([^>]+)>')
        if ($smtpMatch.Success) {
            $result = $smtpMatch.Groups[1].Value.Trim()
        } else {
            $result = $PassedCN
        }
        Write-Verbose "GetSMTPAddress: Using <SMTPAddress> format of [$PassedCN] as [$result]"
    } elseif ($PassedCN -match '</O=') {
        #Matching "Users Name" </O=...>
        $pattern = "<([^>]*)>"
        $MailboxOU = ([regex]::Matches($PassedCN, $pattern)).groups[1].value
        Write-Verbose "GetSMTPAddress: Using /OU format to look up mailbox for [$MailboxOU]"
        $result = GetMailboxProp -PassedCN $MailboxOU -Prop "PrimarySmtpAddress"
    } elseif ($PassedCN -match 'cn=([\w,\s.@-]*[^/])$') {
        Write-Verbose "GetSMTPAddress: Using CN format to look up mailbox for [$PassedCN]"
        $result = GetMailboxProp -PassedCN $PassedCN -Prop "PrimarySmtpAddress"
    } elseif (($PassedCN -match "NotFound") -or ([string]::IsNullOrEmpty($PassedCN))) {
        $result = ""
    } elseif ($PassedCN -match "InvalidSchemaPropertyName") {
        Write-Verbose "GetSMTPAddress: Passed in Value is empty or not Valid: [$PassedCN]."
        $result = ""
    } elseif ($PassedCN -match "@") {
        Write-Verbose "GetSMTPAddress: Looks like we have an SMTP Address already: [$PassedCN]."
        $result = $PassedCN
    } else {
        # We have a problem, we don't have a CN or an SMTP Address
        Write-Warning "GetSMTPAddress: Passed in Value does not look like a CN or SMTP Address: [$PassedCN]."
        $result = $PassedCN
    }

    $script:SMTPAddressCache[$cacheKey] = $result
    return $result
}

<#
.SYNOPSIS
    This function gets a more readable Name from a CN or the Calendar Assistant.
.PARAMETER PassedCN
    The common name (CN) of the mailbox user or the Calendar Assistant.
.OUTPUTS
    Returns the last part of the CN so that it is more readable
#>
function BetterThanNothingCNConversion {
    param (
        $PassedCN
    )
    if ($PassedCN -match $WellKnownCN_CA) {
        return $CalAttendant
    }

    if ($PassedCN -match $WellKnownCN_Trans) {
        return $Transport
    }

    if ($PassedCN -match 'cn=([\w,\s.@-]*[^/])$') {
        $cNameMatch = $PassedCN -split "cn="

        # Normally a readable name is sectioned off with a "-" at the end.
        # Example /o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=d61149258ba04404adda42f336b504ed-Delegate
        if ($cNameMatch[-1] -match "-[\w* -.]*") {
            Write-Verbose "BetterThanNothingCNConversion: Matched : [$($cNameMatch[-1])]"
            $cNameSplit = $cNameMatch.split('-')[-1]
            # Sometimes we have a more than one "-" in the name, so we end up with only 1-4 chars which is too little.
            # Example: .../CN=RECIPIENTS/CN=83DAA772E6A94DA19402AA6B41770486-4DB5F0EB-4A
            if ($cNameSplit.length -lt 5) {
                Write-Verbose "BetterThanNothingCNConversion: [$cNameSplit] is too short"
                $cNameSplit= $cNameMatch.split('-')[-2] + '-' + $cNameMatch.split('-')[-1]
                Write-Verbose "BetterThanNothingCNConversion: Returning Lengthened : [$cNameSplit]"
            }
            Write-Verbose "BetterThanNothingCNConversion: Returning : [$cNameSplit]."
            return $cNameSplit
        }
        # Sometimes we do not have the "-" in front of the Name.
        # Example: "/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=user123"
        if ($cNameMatch[-1] -match "[\w* -.]*") {
            Write-Verbose "BetterThanNothingCNConversion: Returning : [$($cNameMatch[-1])]"
            return $cNameMatch.split('-')[-1]
        }
    }
}

<#
.SYNOPSIS
Checks if an entries is Redacted to protect PII.
#>
function CheckForPIIAccess {
    param(
        $PassedString
    )
    if ($PassedString -match "REDACTED-") {
        return $false
    } else {
        return $true
    }
}

<#
.SYNOPSIS
    Retrieves mailbox properties for a given mailbox.
.DESCRIPTION
    This function retrieves mailbox properties for a given mailbox using Exchange Web Services (EWS).
.PARAMETER CN of the Mailbox
    The mailbox for which to retrieve properties.
.PARAMETER PropertySet
    The set of properties to retrieve.
#>
function GetMailboxProp {
    param(
        $PassedCN,
        $Prop
    )

    Write-Debug "GetMailboxProp: [$Prop]: Searching for:[$PassedCN]..."

    if (($Prop -ne "PrimarySmtpAddress") -and ($Prop -ne "DisplayName")) {
        Write-Error "GetMailboxProp:Invalid Property: [$Prop]"
        return "Invalid Property"
    }

    if ($script:MailboxList.count -gt 0) {
        switch -Regex ($PassedCN) {
            $WellKnownCN_CA {
                return $CalAttendant
            }
            $WellKnownCN_Trans {
                return $Transport
            }
            default {
                if ($null -ne $script:MailboxList[$PassedCN]) {
                    $ReturnValue = $script:MailboxList[$PassedCN].$Prop

                    if ($null -eq $ReturnValue) {
                        Write-Error "`t GetMailboxProp:$Prop :NotFound for ::[$PassedCN]"
                        return BetterThanNothingCNConversion($PassedCN)
                    }

                    Write-Verbose "`t GetMailboxProp:[$Prop] :Found::[$ReturnValue]"
                    if (-not (CheckForPIIAccess($ReturnValue))) {
                        Write-Verbose "No PII Access for [$ReturnValue]"
                        return BetterThanNothingCNConversion($PassedCN)
                    }
                    return $ReturnValue
                } else {
                    Write-Verbose "`t GetMailboxProp:$Prop :NotFound::$PassedCN"
                    return BetterThanNothingCNConversion($PassedCN)
                }
            }
        }
    } else {
        Write-Host -ForegroundColor Red "$script:MailboxList is empty, unable to do CN to SMTP mapping."
        return BetterThanNothingCNConversion($PassedCN)
    }
}

# ===================================================================================================
# Constants to support the script
# ===================================================================================================

$script:CustomPropertyNameList =
"AppointmentCounterProposal",
"AppointmentLastSequenceNumber",
"AppointmentRecurring",
"CalendarItemType",
"CalendarLogTriggerAction",
"CalendarProcessed",
"ChangeList",
"ClientBuildVersion",
"ClientIntent",
"ClientProcessName",
"CreationTime",
"DisplayAttendeesCc",
"DisplayAttendeesTo",
"EventEmailReminderTimer",
"ExternalSharingMasterId",
"FreeBusyStatus",
"From",
"HasAttachment",
"InternetMessageId",
"IsAllDayEvent",
"IsCancelled",
"IsException",
"IsMeeting",
"IsOrganizerProperty",
"IsSharedInEvent",
"ItemID",
"LogBodyStats",
"LogClientInfoString",
"LogRowType",
"LogTimestamp",
"NormalizedSubject",
"OriginalStartDate",
"ReminderDueByInternal",
"ReminderIsSetInternal",
"ReminderMinutesBeforeStartInternal",
"SendMeetingMessagesDiagnostics",
"Sensitivity",
"SentRepresentingDisplayName",
"ShortClientInfoString",
"TimeZone"

$LogLimit = 2000

if ($ShortLogs.IsPresent) {
    $LogLimit = 500
}

if ($MaxLogs.IsPresent) {
    $LogLimit = 12000
}

$LimitedItemClasses = @(
    "IPM.Appointment",
    "IPM.Schedule.Meeting.Request",
    "IPM.Schedule.Meeting.Canceled",
    "IPM.Schedule.Meeting.Forwarded"
)

<#
.SYNOPSIS
Run Get-CalendarDiagnosticObjects for passed in User with Subject or MeetingID.
#>
function GetCalendarDiagnosticObjects {
    param(
        [string]$Identity,
        [string]$Subject,
        [string]$MeetingID,
        [datetime]$ExceptionDateOverride
    )

    $params = @{
        Identity           = $Identity
        CustomPropertyName = $script:CustomPropertyNameList
        WarningAction      = "Ignore"
        MaxResults         = $LogLimit
        ResultSize         = $LogLimit
        ShouldBindToItem   = $true
        ShouldDecodeEnums  = $true
    }

    $isExceptionOverrideCall = $PSBoundParameters.ContainsKey('ExceptionDateOverride') -and $null -ne $ExceptionDateOverride

    if ($TrackingLogs -eq $true) {
        if (-not $isExceptionOverrideCall) {
            Write-Host -ForegroundColor Yellow "Including Tracking Logs in the output."
        }
        $script:CustomPropertyNameList += "AttendeeListDetails", "AttendeeCollection"
        $params.Add("ShouldFetchAttendeeCollection", $true)
        $params.Remove("CustomPropertyName")
        $params.Add("CustomPropertyName", $script:CustomPropertyNameList)
    }

    $effectiveExceptionDate = if ($PSBoundParameters.ContainsKey('ExceptionDateOverride') -and $null -ne $ExceptionDateOverride) { $ExceptionDateOverride } else { $ExceptionDate }

    if (-not [string]::IsNullOrEmpty($effectiveExceptionDate)) {
        Write-Host -ForegroundColor Green "---------------------------------------"
        $exceptionDateLabel = if ($effectiveExceptionDate -is [datetime]) { $effectiveExceptionDate.ToString('MM/dd/yyyy') } else { [string]$effectiveExceptionDate }
        Write-Host -ForegroundColor Green "Pulling all the Exceptions for [$exceptionDateLabel] and adding them to the output."
        Write-Host -ForegroundColor Green "---------------------------------------"
        $params.Add("AnalyzeExceptionWithOriginalStartDate", $effectiveExceptionDate)
    }

    if ($MaxLogs.IsPresent) {
        Write-Host -ForegroundColor Yellow "Limiting the number of logs to $LogLimit, and limiting the number of Item Classes retrieved."
        $params.Add("ItemClass", $LimitedItemClasses)
    }

    if ($null -ne $CustomProperty) {
        Write-Host -ForegroundColor Yellow "Adding custom properties to the RAW output."
        $params.Remove("CustomPropertyName")
        $script:CustomPropertyNameList += $CustomProperty
        Write-Host -ForegroundColor Yellow "Adding extra CustomProperty: [$CustomProperty]"
        $params.Add("CustomPropertyName", $script:CustomPropertyNameList)
    }

    # Use 3>$null to suppress "Non-view properties may not be accurate for exception items"
    # warnings from Get-CalendarDiagnosticObjects (remote session warnings bypass WarningAction/Preference).
    if ($Identity -and $MeetingID) {
        Write-Verbose "Getting CalLogs for [$Identity] with MeetingID [$MeetingID]."
        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {
            Write-Host -ForegroundColor Yellow ($params.GetEnumerator() | ForEach-Object { "`t$($_.Key) = $($_.Value)`n" })
        }
        $CalLogs = Get-CalendarDiagnosticObjects @params -MeetingID $MeetingID 3>$null
    } elseif ($Identity -and $Subject ) {
        Write-Verbose "Getting CalLogs for [$Identity] with Subject [$Subject]."
        $CalLogs = Get-CalendarDiagnosticObjects @params -Subject $Subject 3>$null

        # No Results, do a Deep search with ExactMatch.
        if ($CalLogs.count -lt 1) {
            $CalLogs = Get-CalendarDiagnosticObjects @Params -Subject $Subject -ExactMatch $true 3>$null
        }
    }

    if (-not $isExceptionOverrideCall) {
        Write-Host "Found $($CalLogs.count) Calendar Logs for [$Identity]"
    }
    return $CalLogs
}

<#
.SYNOPSIS
This function retrieves calendar logs from the specified source with a subject that matches the provided criteria.
.PARAMETER Identity
The Identity of the mailbox to get calendar logs from.
.PARAMETER Subject
The subject of the calendar logs to retrieve.
#>
function GetCalLogsWithSubject {
    param (
        [string] $Identity,
        [string] $Subject
    )
    Write-Host "Getting CalLogs from [$Identity] with subject [$Subject]"
    $script:CurrentIdentityRunStartTime = Get-Date

    $InitialCDOs = GetCalendarDiagnosticObjects -Identity $Identity -Subject $Subject

    # Find all the unique Global Object IDs
    $GlobalObjectIds = @($InitialCDOs.CleanGlobalObjectId | Where-Object {
            -not [string]::IsNullOrEmpty($_) -and
            $_ -ne "NotFound" -and
            $_ -ne "InvalidSchemaPropertyName" -and
            $_.Length -ge 90
        } | Select-Object -Unique)
    $script:SubjectMeetingIdCount = $GlobalObjectIds.Count
    $script:SubjectResolvedMeetingId = $null
    $script:SubjectCanCollectExceptions = $false
    $script:SubjectSkippedExceptionCollection = $false
    Write-Host "Found $($GlobalObjectIds.count) unique GlobalObjectIds."
    Write-Host "Getting the set of CalLogs for each GlobalObjectID."

    if ($GlobalObjectIds.count -eq 1) {
        $script:SubjectResolvedMeetingId = $GlobalObjectIds[0]
        $script:SubjectCanCollectExceptions = $Exceptions.IsPresent -and [string]::IsNullOrEmpty($ExceptionDate)
        $script:GCDO = $InitialCDOs; # use the CalLogs that we already have, since there is only one.
        if ($Exceptions.IsPresent -and [string]::IsNullOrEmpty($ExceptionDate)) {
            Write-Host -ForegroundColor Yellow "Subject search resolved to one MeetingID [$($script:SubjectResolvedMeetingId)]. Collecting Exceptions by default."
            CollectExceptionLogs -Identity $Identity -MeetingID $script:SubjectResolvedMeetingId
        } elseif (-not [string]::IsNullOrEmpty($ExceptionDate)) {
            $script:ExceptionCollectionStatus = "SkippedUntilMeetingIdChosen"
            Write-Host -ForegroundColor Yellow "Subject search resolved to one MeetingID [$($script:SubjectResolvedMeetingId)], but -ExceptionDate requires rerunning with -MeetingID for targeted Exception collection."
        } else {
            $script:ExceptionCollectionStatus = "SkippedBySwitch"
            Write-Host -ForegroundColor Green "Subject search resolved to one MeetingID [$($script:SubjectResolvedMeetingId)], but Exception collection was skipped."
        }
        BuildCSV
        BuildTimeline
    } elseif ($GlobalObjectIds.count -gt 1) {
        $script:SubjectSkippedExceptionCollection = $true
        $script:ExceptionCollectionStatus = "SkippedMultipleMeetingIds"
        Write-Host "Found multiple GlobalObjectIds: $($GlobalObjectIds.Count)."
        Write-Host -ForegroundColor Yellow "Exception collection is skipped when Subject search resolves to more than one MeetingID."
        Write-Host -ForegroundColor Yellow "Re-run with one of these MeetingIDs to collect meeting exceptions:"
        foreach ($MID in $GlobalObjectIds) {
            Write-Host -ForegroundColor Yellow "  -MeetingID $MID"
        }
        foreach ($MID in $GlobalObjectIds) {
            $script:CurrentIdentityRunStartTime = Get-Date
            Write-DashLineBoxColor "Processing MeetingID: [$MID]"
            $script:GCDO = GetCalendarDiagnosticObjects -Identity $Identity -MeetingID $MID
            Write-Verbose "Found $($GCDO.count) CalLogs with MeetingID[$MID] ."
            BuildCSV
            BuildTimeline
        }
    } else {
        $script:ExceptionCollectionStatus = "NoMeetingId"
        Write-Warning "No CalLogs were found."
    }
}

<#
.SYNOPSIS
Checks if a set of Calendar Logs is from the Organizer.
#>
function SetIsOrganizer {
    param(
        $CalLogs
    )
    [bool] $IsOrganizer = $false

    foreach ($CalLog in $CalLogs) {
        if ($CalLog.ItemClass -eq "Ipm.Appointment" -and
            $CalLog.ExternalSharingMasterId -eq "NotFound" -and
            ($CalLog.ResponseType -eq "1" -or $CalLog.ResponseType -eq "Organizer")) {
            $IsOrganizer = $true
            Write-Host -ForegroundColor Green "IsOrganizer: [$IsOrganizer]"
            return $IsOrganizer
        }
    }
    Write-Verbose "IsOrganizer: [$IsOrganizer]"
    return $IsOrganizer
}

<#
.SYNOPSIS
Checks if a set of Calendar Logs is from a Resource Mailbox.
#>
function SetIsRoom {
    param(
        $CalLogs
    )

    # See if we have already determined this is a Room MB.
    if ($script:Rooms -contains $Identity) {
        return $true
    }

    # Simple logic is if RBA is running on the MB, it is a Room MB, otherwise it is not.
    $rbaLog = $CalLogs | Where-Object {
        $_.ItemClass -eq "IPM.Appointment" -and
        $_.ExternalSharingMasterId -eq "NotFound" -and
        $_.LogClientInfoString -like "*ResourceBookingAssistant*"
    } | Select-Object -First 1

    if ($null -ne $rbaLog) {
        Write-Host -ForegroundColor Green "Found Room Mailbox indicator: [$($rbaLog.LogClientInfoString)]"
        return $true
    }
    return $false
}

<#
.SYNOPSIS
Checks if a set of Calendar Logs is from a Recurring Meeting.
#>
function SetIsRecurring {
    param(
        $CalLogs
    )
    Write-Host -ForegroundColor Yellow "Looking for signs of a recurring meeting."
    [bool] $IsRecurring = $false
    # See if this is a recurring meeting
    foreach ($CalLog in $CalLogs) {
        if ($null -ne $CalLog.ItemClass -and
            (GetItemType $CalLog.ItemClass) -eq "Ipm.Appointment" -and
            # Commenting this out will get all the updates for shared calendars, which is important with Delegates.
            #      $CalLog.ExternalSharingMasterId -eq "NotFound" -and
            $CalLog.CalendarItemType.ToString() -eq "RecurringMaster") {
            $IsRecurring = $true
            Write-Verbose "Found recurring meeting."
            return $IsRecurring
        }
    }
    Write-Verbose "Did not find signs of recurring meeting."
    return $IsRecurring
}

<#
.SYNOPSIS
Check for Bifurcation issue
#>
function CheckForBifurcation {
    param (
        $CalLogs
    )
    Write-Verbose  "Looking for signs of the Bifurcation Issue."
    [bool] $IsBifurcated = $false
    # See if there is an IPM.Appointment in the CalLogs.
    foreach ($CalLog in $CalLogs) {
        if ($CalLog.ItemClass -eq "IPM.Appointment" -and
            $CalLog.ExternalSharingMasterId -eq "NotFound") {
            $IsBifurcated = $false
            Write-Verbose "Found Ipm.Appointment, likely not a bifurcation issue."
            return $IsBifurcated
        }
    }
    Write-Host -ForegroundColor Red "Did not find any Ipm.Appointments in the CalLogs. If this is the Organizer of the meeting, this could be the Outlook Bifurcation issue."
    Write-Host -ForegroundColor Yellow "`t The Outlook Bifurcation issue is where Outlook saves to the Organizer's Mailbox on one thread and sends to the attendees via transport on another thread. If the save to the Organizer's mailbox failed, we get into the Bifurcated State, where the Organizer does not have the meeting but the Attendees do."
    Write-Host -ForegroundColor Yellow "`t See https://support.microsoft.com/en-us/office/meeting-request-is-missing-from-organizers-calendar-c13c47cd-18f9-4ef0-b9d0-d9e174912c4a"
    return $IsBifurcated
}

function Read-UInt16LE {
    param(
        [byte[]]$Bytes,
        [ref]$Offset
    )

    if (($Offset.Value + 2) -gt $Bytes.Length) {
        throw "Unexpected end of AppointmentRecurrenceBlob while reading UInt16 at offset [$($Offset.Value)]."
    }

    $value = [BitConverter]::ToUInt16($Bytes, $Offset.Value)
    $Offset.Value += 2
    return $value
}

function Read-UInt32LE {
    param(
        [byte[]]$Bytes,
        [ref]$Offset
    )

    if (($Offset.Value + 4) -gt $Bytes.Length) {
        throw "Unexpected end of AppointmentRecurrenceBlob while reading UInt32 at offset [$($Offset.Value)]."
    }

    $value = [BitConverter]::ToUInt32($Bytes, $Offset.Value)
    $Offset.Value += 4
    return $value
}

function Convert-RecurMinutesToDateTime {
    param(
        [UInt32]$Minutes
    )

    return ([datetime]'1601-01-01').AddMinutes([double]$Minutes)
}

function Get-RecurrencePatternSpecificSize {
    param(
        [UInt16]$PatternType
    )

    switch ($PatternType) {
        0 { return 0 }
        1 { return 4 }
        2 { return 4 }
        3 { return 8 }
        4 { return 4 }
        10 { return 4 }
        11 { return 8 }
        12 { return 4 }
        default {
            throw "Unsupported recurrence PatternType [$PatternType] in AppointmentRecurrenceBlob."
        }
    }
}

function Convert-AppointmentRecurrenceBlobToBytes {
    param(
        $AppointmentRecurrenceBlob
    )

    if ($AppointmentRecurrenceBlob -is [byte[]]) {
        return $AppointmentRecurrenceBlob
    }

    $blobText = [string]$AppointmentRecurrenceBlob
    if ([string]::IsNullOrWhiteSpace($blobText)) {
        throw "AppointmentRecurrenceBlob is empty."
    }

    $blobText = ($blobText -replace '\s+', '').Trim()
    if (($blobText.Length % 2) -ne 0) {
        throw "AppointmentRecurrenceBlob length is invalid. Expected an even-length hex string."
    }

    $bytes = [byte[]]::new($blobText.Length / 2)
    for ($index = 0; $index -lt $blobText.Length; $index += 2) {
        $bytes[$index / 2] = [Convert]::ToByte($blobText.Substring($index, 2), 16)
    }

    return $bytes
}

function Get-RecentExceptionCutoff {
    if ($FastExceptions.IsPresent -and -not $AllExceptions.IsPresent) {
        return (Get-Date).Date.AddMonths(-6)
    }

    return $null
}

function Get-CalendarDiagnosticObjectDeduplicationKey {
    param(
        $CalLog
    )

    $itemId = ""
    if ($null -ne $CalLog.ItemId) {
        if ($CalLog.ItemId.PSObject.Properties.Name -contains 'ObjectId') {
            $itemId = [string]$CalLog.ItemId.ObjectId
        } else {
            $itemId = [string]$CalLog.ItemId
        }
    }

    $calendarLogRequestId = if ($null -ne $CalLog.CalendarLogRequestId) { [string]$CalLog.CalendarLogRequestId } else { "" }
    $logTimestamp = if ($null -ne $CalLog.LogTimestamp) { [string]$CalLog.LogTimestamp } else { "" }
    $originalStartDate = if ($null -ne $CalLog.OriginalStartDate) { [string]$CalLog.OriginalStartDate } else { "" }
    $triggerAction = if ($null -ne $CalLog.CalendarLogTriggerAction) { [string]$CalLog.CalendarLogTriggerAction } else { "" }
    $itemClass = if ($null -ne $CalLog.ItemClass) { [string]$CalLog.ItemClass } else { "" }
    $itemVersion = if ($null -ne $CalLog.ItemVersion) { [string]$CalLog.ItemVersion } else { "" }
    $logRowType = if ($null -ne $CalLog.LogRowType) { [string]$CalLog.LogRowType } else { "" }
    $responsibleUserName = if ($null -ne $CalLog.ResponsibleUserName) { [string]$CalLog.ResponsibleUserName } else { "" }
    $clientInfo = if ($null -ne $CalLog.LogClientInfoString) { [string]$CalLog.LogClientInfoString } else { "" }

    return "$itemId|$calendarLogRequestId|$logTimestamp|$originalStartDate|$triggerAction|$itemClass|$itemVersion|$logRowType|$responsibleUserName|$clientInfo"
}

function RemoveDuplicateCalendarDiagnosticObjects {
    param(
        [array]$CalLogs,
        [switch]$Quiet
    )

    if ($null -eq $CalLogs -or $CalLogs.Count -le 1) {
        return @($CalLogs)
    }

    $uniqueLogs = @($CalLogs | Group-Object -Property { Get-CalendarDiagnosticObjectDeduplicationKey -CalLog $_ } | ForEach-Object {
            $_.Group | Select-Object -First 1
        } | Sort-Object { ConvertDateTime($_.LogTimestamp.ToString()) })

    $duplicateCount = $CalLogs.Count - $uniqueLogs.Count
    if (($duplicateCount -gt 0) -and (-not $Quiet.IsPresent)) {
        Write-Host -ForegroundColor Yellow "Removed [$duplicateCount] duplicate Calendar Log entries before processing."
    }

    return $uniqueLogs
}

function Merge-CalendarDiagnosticObjects {
    param(
        [array]$BaseLogs,
        [array]$AdditionalLogs
    )

    $combinedLogs = @($BaseLogs) + @($AdditionalLogs)
    return RemoveDuplicateCalendarDiagnosticObjects -CalLogs $combinedLogs -Quiet
}

function FilterExceptionLogsByRecency {
    param(
        [array]$ExceptionLogs
    )

    $cutoffDate = Get-RecentExceptionCutoff
    if ($null -eq $cutoffDate) {
        return @($ExceptionLogs)
    }

    return @($ExceptionLogs | Where-Object {
            if ($null -eq $_.OriginalStartDate -or [string]::IsNullOrEmpty([string]$_.OriginalStartDate) -or $_.OriginalStartDate -eq "NotFound") {
                return $false
            }

            $originalStartDate = ConvertDateTime($_.OriginalStartDate.ToString())
            return ($originalStartDate -is [datetime]) -and $originalStartDate -ne [datetime]::MinValue -and $originalStartDate -ge $cutoffDate
        })
}

function Get-AppointmentExceptionDatesFromBlob {
    param(
        $AppointmentRecurrenceBlob
    )

    [byte[]]$bytes = Convert-AppointmentRecurrenceBlobToBytes -AppointmentRecurrenceBlob $AppointmentRecurrenceBlob
    $offset = [ref]0

    $readerVersion = Read-UInt16LE -Bytes $bytes -Offset $offset
    $writerVersion = Read-UInt16LE -Bytes $bytes -Offset $offset
    $null = Read-UInt16LE -Bytes $bytes -Offset $offset
    $patternType = Read-UInt16LE -Bytes $bytes -Offset $offset
    $null = Read-UInt16LE -Bytes $bytes -Offset $offset
    $null = Read-UInt32LE -Bytes $bytes -Offset $offset
    $null = Read-UInt32LE -Bytes $bytes -Offset $offset
    $null = Read-UInt32LE -Bytes $bytes -Offset $offset

    $patternSpecificSize = Get-RecurrencePatternSpecificSize -PatternType $patternType
    if (($offset.Value + $patternSpecificSize) -gt $bytes.Length) {
        throw "AppointmentRecurrenceBlob ended before PatternTypeSpecific data was fully read."
    }
    $offset.Value += $patternSpecificSize

    $null = Read-UInt32LE -Bytes $bytes -Offset $offset
    $null = Read-UInt32LE -Bytes $bytes -Offset $offset
    $null = Read-UInt32LE -Bytes $bytes -Offset $offset

    $deletedInstanceCount = Read-UInt32LE -Bytes $bytes -Offset $offset
    $deletedInstanceDates = [System.Collections.Generic.List[datetime]]::new()
    for ($index = 0; $index -lt $deletedInstanceCount; $index++) {
        $deletedInstanceDates.Add((Convert-RecurMinutesToDateTime -Minutes (Read-UInt32LE -Bytes $bytes -Offset $offset)).Date)
    }

    $modifiedInstanceCount = Read-UInt32LE -Bytes $bytes -Offset $offset
    $modifiedInstanceDates = [System.Collections.Generic.List[datetime]]::new()
    for ($index = 0; $index -lt $modifiedInstanceCount; $index++) {
        $modifiedInstanceDates.Add((Convert-RecurMinutesToDateTime -Minutes (Read-UInt32LE -Bytes $bytes -Offset $offset)).Date)
    }

    $seriesStartDate = Convert-RecurMinutesToDateTime -Minutes (Read-UInt32LE -Bytes $bytes -Offset $offset)
    $seriesEndDate = Convert-RecurMinutesToDateTime -Minutes (Read-UInt32LE -Bytes $bytes -Offset $offset)
    $readerVersion2 = Read-UInt32LE -Bytes $bytes -Offset $offset
    $writerVersion2 = Read-UInt32LE -Bytes $bytes -Offset $offset
    $null = Read-UInt32LE -Bytes $bytes -Offset $offset
    $null = Read-UInt32LE -Bytes $bytes -Offset $offset
    $exceptionCount = Read-UInt16LE -Bytes $bytes -Offset $offset

    if ($readerVersion -ne 0x3004 -or $writerVersion -ne 0x3004) {
        throw "Unexpected recurrence blob header versions [$readerVersion/$writerVersion]."
    }
    if ($readerVersion2 -lt 0x3006 -or $writerVersion2 -lt 0x3006) {
        throw "Unexpected recurrence blob secondary versions [$readerVersion2/$writerVersion2]."
    }
    if ($modifiedInstanceCount -lt $exceptionCount) {
        throw "ModifiedInstanceCount [$modifiedInstanceCount] is less than ExceptionCount [$exceptionCount]."
    }

    return [PSCustomObject]@{
        DeletedInstanceCount  = $deletedInstanceCount
        ModifiedInstanceCount = $modifiedInstanceCount
        ExceptionCount        = $exceptionCount
        DeletedInstanceDates  = @($deletedInstanceDates)
        ModifiedInstanceDates = @($modifiedInstanceDates)
        Dates                 = @($deletedInstanceDates + $modifiedInstanceDates | Sort-Object -Unique)
        SeriesStartDate       = $seriesStartDate
        SeriesEndDate         = $seriesEndDate
        PatternType           = $patternType
    }
}

function Get-ExceptionDatesFromMostRecentAppointmentBlob {
    param(
        [array]$CalLogs
    )

    $recurringAppointments = @($CalLogs | Where-Object {
            $_.ItemClass -eq 'IPM.Appointment' -and
            $_.AppointmentRecurring -and
            $null -ne $_.AppointmentRecurrenceBlob -and
            -not [string]::IsNullOrWhiteSpace([string]$_.AppointmentRecurrenceBlob)
        })

    if ($recurringAppointments.Count -eq 0) {
        throw "No recurring IPM.Appointment with an AppointmentRecurrenceBlob was found in the Calendar Logs."
    }

    $blobSource = $recurringAppointments | Where-Object { $_.CalendarItemType.ToString() -eq 'RecurringMaster' } | Sort-Object @{ Expression = { $itemVersion = 0; [void][int]::TryParse([string]$_.ItemVersion, [ref]$itemVersion); $itemVersion } }, @{ Expression = { ConvertDateTime($_.LogTimestamp.ToString()) } } -Descending | Select-Object -First 1
    if ($null -eq $blobSource) {
        $blobSource = $recurringAppointments | Sort-Object @{ Expression = { $itemVersion = 0; [void][int]::TryParse([string]$_.ItemVersion, [ref]$itemVersion); $itemVersion } }, @{ Expression = { ConvertDateTime($_.LogTimestamp.ToString()) } } -Descending | Select-Object -First 1
    }

    $parsedBlob = Get-AppointmentExceptionDatesFromBlob -AppointmentRecurrenceBlob $blobSource.AppointmentRecurrenceBlob
    $exceptionDates = @($parsedBlob.Dates)

    $cutoffDate = Get-RecentExceptionCutoff
    if ($null -ne $cutoffDate) {
        $exceptionDates = @($exceptionDates | Where-Object { $_ -ge $cutoffDate })
    }

    return [PSCustomObject]@{
        SourceLog      = $blobSource
        ParsedBlob     = $parsedBlob
        ExceptionDates = $exceptionDates
    }
}

function CollectExceptionLogsLegacy {
    param(
        [string]$Identity,
        [string]$MeetingID,
        [switch]$UsedAsFallback
    )

    $ExceptionLogs = @()
    $LogToExamine = @()
    $LogToExamine = $script:GCDO | Where-Object { $_.ItemClass -like 'IPM.Appointment*' } | Sort-Object ItemVersion

    Write-Host -ForegroundColor Cyan "Found $($LogToExamine.count) CalLogs to examine for Exception Logs."
    if ($LogToExamine.count -gt 100) {
        Write-Host -ForegroundColor Cyan "`t This is a large number of logs to examine, this may take a while."
    }
    $logLeftCount = $LogToExamine.count

    $ExceptionLogs = $LogToExamine | ForEach-Object {
        $logLeftCount -= 1
        Write-Verbose "Getting Exception Logs for [$($_.ItemId.ObjectId)]"
        Get-CalendarDiagnosticObjects -Identity $Identity -ItemIds $_.ItemId.ObjectId -ShouldFetchRecurrenceExceptions $true -CustomPropertyNames $CustomPropertyNameList -ShouldBindToItem $true 3>$null
        if (($logLeftCount % 10 -eq 0) -and ($logLeftCount -gt 0)) {
            Write-Host -ForegroundColor Cyan "`t [$($logLeftCount)] logs left to examine..."
        }
    }

    $ExceptionLogs = $ExceptionLogs | Where-Object { $_.ItemClass -notlike "IPM.Appointment*" }
    $cutoffDate = Get-RecentExceptionCutoff
    if ($null -ne $cutoffDate) {
        Write-Host -ForegroundColor Yellow "Filtering legacy Exception logs to only keep items with OriginalStartDate in the last 6 months."
        $ExceptionLogs = FilterExceptionLogsByRecency -ExceptionLogs $ExceptionLogs
    }

    Write-Host -ForegroundColor Cyan "Found $($ExceptionLogs.count) Exception Logs, adding them into the CalLogs."
    $script:GCDO = Merge-CalendarDiagnosticObjects -BaseLogs $script:GCDO -AdditionalLogs $ExceptionLogs
    if ($UsedAsFallback.IsPresent) {
        $script:ExceptionCollectionStatus = "CollectedLegacyFallback"
    } else {
        $script:ExceptionCollectionStatus = "Collected"
    }
}

function CollectExceptionLogsFast {
    param(
        [string]$Identity,
        [string]$MeetingID
    )

    $blobResult = Get-ExceptionDatesFromMostRecentAppointmentBlob -CalLogs $script:GCDO
    $exceptionDates = @($blobResult.ExceptionDates | Sort-Object -Unique)

    Write-Host -ForegroundColor Cyan "Fast exception collection selected blob from ItemVersion [$($blobResult.SourceLog.ItemVersion)] with [$($blobResult.ParsedBlob.ExceptionCount)] exception entries."
    Write-Host -ForegroundColor Cyan "Found [$($blobResult.ParsedBlob.DeletedInstanceCount)] deleted exception dates and [$($blobResult.ParsedBlob.ModifiedInstanceCount)] modified exception dates in the recurrence blob."

    $cutoffDate = Get-RecentExceptionCutoff
    if ($null -ne $cutoffDate) {
        Write-Host -ForegroundColor Cyan "Keeping only Exception dates from the last 6 months: [$($cutoffDate.ToString('yyyy-MM-dd'))] and newer."
        Write-Host -ForegroundColor Cyan "[$($exceptionDates.Count)] Exception date(s) match the 6 month time frame."
    }

    if ($exceptionDates.Count -eq 0) {
        $script:ExceptionCollectionStatus = "NoExceptionDates"
        Write-Host -ForegroundColor Cyan "No matching Exception dates were found in the AppointmentRecurrenceBlob."
        return
    }

    $collectedExceptionLogs = @(
        foreach ($exceptionOriginalStartDate in $exceptionDates) {
            GetCalendarDiagnosticObjects -Identity $Identity -MeetingID $MeetingID -ExceptionDateOverride $exceptionOriginalStartDate
        }
    )

    $collectedExceptionLogs = @($collectedExceptionLogs | Where-Object {
            $_.ItemClass -notlike "IPM.Appointment*" -or
            ($null -ne $_.OriginalStartDate -and $_.OriginalStartDate -ne "NotFound" -and -not [string]::IsNullOrEmpty([string]$_.OriginalStartDate))
        })

    Write-Host -ForegroundColor Cyan "Collected $($collectedExceptionLogs.count) Exception-related logs from [$($exceptionDates.count)] ExceptionDate queries."
    $script:GCDO = Merge-CalendarDiagnosticObjects -BaseLogs $script:GCDO -AdditionalLogs $collectedExceptionLogs
    $script:ExceptionCollectionStatus = "CollectedFast"
}

function CollectExceptionLogs {
    param(
        [string]$Identity,
        [string]$MeetingID
    )

    Write-Verbose "Looking for Exception Logs..."
    $IsRecurring = SetIsRecurring -CalLogs $script:GCDO
    Write-Verbose "Meeting IsRecurring: $IsRecurring"

    if ($IsRecurring) {
        if ($FastExceptions.IsPresent -and [string]::IsNullOrEmpty($ExceptionDate)) {
            try {
                CollectExceptionLogsFast -Identity $Identity -MeetingID $MeetingID
            } catch {
                $script:ExceptionCollectionStatus = "CollectedLegacyFallback"
                Write-DashLineBoxColor "FAST EXCEPTION COLLECTION FAILED",
                "Error: $($_.Exception.Message)",
                "Falling back to the legacy per-appointment Exception collector." -Color Red -DashChar "="
                CollectExceptionLogsLegacy -Identity $Identity -MeetingID $MeetingID -UsedAsFallback
            }
        } else {
            CollectExceptionLogsLegacy -Identity $Identity -MeetingID $MeetingID
        }
    } else {
        $script:ExceptionCollectionStatus = "NotRecurring"
        Write-Host -ForegroundColor Cyan "No Recurring Meetings found, no Exception Logs to collect."
    }
}

# ===================================================================================================
# FileNames
# ===================================================================================================
function Get-FileName {
    Write-Host -ForegroundColor Cyan "Creating FileName for $Identity..."

    $ThisMeetingID = $script:GCDO.CleanGlobalObjectId | Select-Object -Unique
    $ShortMeetingID = $ThisMeetingID.Substring($ThisMeetingID.length - 6)

    if ($script:Identity -like "*@*") {
        $script:ShortId = $script:Identity.split('@')[0]
    } else {
        $script:ShortId = $script:Identity
    }
    $script:ShortId = $ShortId.Substring(0, [System.Math]::Min(20, $ShortId.Length))

    if (($null -eq $CaseNumber) -or
        ([string]::IsNullOrEmpty($CaseNumber))) {
        $Case = ""
    } else {
        $Case = $CaseNumber + "_"
    }

    if ($ExportToExcel.IsPresent) {
        $script:FileName = "$($Case)CalLogSummary_$($ShortMeetingID).xlsx"
        Write-Host -ForegroundColor Blue -NoNewline "All Calendar Logs for meetings ending in ID [$ShortMeetingID] will be saved to : "
        Write-Host -ForegroundColor Yellow "$Filename"
    } else {
        $script:Filename = "$($Case)$($ShortId)_$ShortMeetingID.csv"
        $script:FilenameRaw = "$($Case)$($ShortId)_RAW_$($ShortMeetingID).csv"
        $Script:TimeLineFilename = "$($Case)$($ShortId)_TimeLine_$ShortMeetingID.csv"

        Write-Host -ForegroundColor Cyan -NoNewline "Enhanced Calendar Logs for [$Identity] has been saved to : "
        Write-Host -ForegroundColor Yellow "$Filename"

        Write-Host -ForegroundColor Cyan -NoNewline "Raw Calendar Logs for [$Identity] has been saved to : "
        Write-Host -ForegroundColor Yellow "$FilenameRaw"

        Write-Host -ForegroundColor Cyan -NoNewline "TimeLine for [$Identity] has been saved to : "
        Write-Host -ForegroundColor Yellow "$TimeLineFilename"
    }
}

function Export-CalLog {
    Get-FileName

    if ($ExportToExcel.IsPresent) {
        Export-CalLogExcel
    } else {
        Export-CalLogCSV
    }
}

function Export-CalLogCSV {
    $script:EnhancedCalLogs | Export-Csv -Path $Filename -NoTypeInformation -Encoding UTF8
    $script:GCDO | Export-Csv -Path $FilenameRaw -NoTypeInformation -Encoding UTF8
}

function Export-Timeline {
    Write-Verbose "Export to Excel is : $ExportToExcel"

    # Display Timeline to screen:
    Write-Host -ForegroundColor Cyan "Timeline for [$Identity]..."
    $script:TimeLineOutput

    if ($ExportToExcel.IsPresent) {
        Export-TimelineExcel
    } else {
        $script:TimeLineOutput | Export-Csv -Path $script:TimeLineFilename -NoTypeInformation -Encoding UTF8 -Append
    }
}

<#
.SYNOPSIS
    This is the part that generates the heart of the timeline, a Giant Switch statement based on the ItemClass.
#>
function CreateTimelineRow {
    switch -Wildcard ($CalLog.ItemClass) {
        Meeting.Request {
            switch ($CalLog.TriggerAction) {
                Create {
                    if ($IsOrganizer) {
                        if ($CalLog.IsException -eq $True) {
                            [array] $Output = "[$($CalLog.ResponsibleUser)] Created an Exception Meeting Request with $($CalLog.Client) for [$($CalLog.StartTime)]."
                        } else {
                            [array] $Output  = "[$($CalLog.ResponsibleUser)] Created a Meeting Request with $($CalLog.Client)"
                        }
                    } else {
                        if ($CalLog.DisplayAttendeesTo -ne $script:PreviousCalLog.DisplayAttendeesTo -or $CalLog.DisplayAttendeesCc -ne $script:PreviousCalLog.DisplayAttendeesCc) {
                            [array] $Output = "The user Forwarded a Meeting Request with $($CalLog.Client)."
                        } else {
                            if ($CalLog.Client -eq "Transport") {
                                if ($CalLog.IsException -eq $True) {
                                    [array] $Output = "Transport delivered a new Meeting Request from [$($CalLog.From)] for an exception starting on [$($CalLog.StartTime)]" + $(if ($null -ne $($CalLog.ReceivedRepresenting)) { " for user [$($CalLog.ReceivedRepresenting)]" }) + "."
                                    $script:MeetingSummaryNeeded = $True
                                } else {
                                    [Array] $Output = "Transport delivered a new Meeting Request from [$($CalLog.From)]" +
                                    $(if ($null -ne $($CalLog.ReceivedRepresenting) -and $CalLog.ReceivedRepresenting -ne $CalLog.ReceivedBy)
                                        { " for user [$($CalLog.ReceivedRepresenting)]" }) + "."
                                }
                            } elseif ($calLog.client -eq "ResourceBookingAssistant") {
                                [array] $Output  = "ResourceBookingAssistant Forwarded a Meeting Request to a Resource Delegate."
                            } elseif ($CalLog.Client -eq "CalendarRepairAssistant") {
                                if ($CalLog.IsException -eq $True) {
                                    [array] $Output = "CalendarRepairAssistant Created a new Meeting Request to repair an inconsistency with an exception starting on [$($CalLog.StartTime)]."
                                } else {
                                    [array] $Output = "CalendarRepairAssistant Created a new Meeting Request to repair an inconsistency."
                                }
                            } else {
                                if ($CalLog.IsException -eq $True) {
                                    [array] $Output = "[$($CalLog.ResponsibleUser)] Created a new Meeting Request with $($CalLog.Client) for an exception starting on [$($CalLog.StartTime)]."
                                } else {
                                    [array] $Output = "[$($CalLog.ResponsibleUser)] Created a new Meeting Request with $($CalLog.Client)."
                                }
                            }
                        }
                    }
                }
                Update {
                    if ($calLog.client -eq "ResourceBookingAssistant") {
                        [array] $Output  = "ResourceBookingAssistant Updated the Meeting Request."
                    } else {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] Updated the $($CalLog.MeetingRequestType.Value) Meeting Request with $($CalLog.Client)."
                    }
                }
                MoveToDeletedItems {
                    if ($CalLog.ResponsibleUser -eq "Calendar Assistant") {
                        [array] $Output = "$($CalLog.Client) Deleted the Meeting Request."
                    } else {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] Deleted the Meeting Request with $($CalLog.Client)."
                    }
                }
                default {
                    [array] $Output = "[$($CalLog.ResponsibleUser)] Deleted the $($CalLog.MeetingRequestType.Value) Meeting Request with $($CalLog.Client)."
                }
            }
        }
        Resp* {
            switch ($CalLog.ItemClass) {
                "Resp.Tent" { $MeetingRespType = "Tentative" }
                "Resp.Neg" { $MeetingRespType = "DECLINE" }
                "Resp.Pos" { $MeetingRespType = "ACCEPT" }
            }

            if ($CalLog.AppointmentCounterProposal -eq "True") {
                [array] $Output = "[$($CalLog.Organizer)] sent a $($MeetingRespType) response message with a New Time Proposal: $($CalLog.StartTime) to $($CalLog.EndTime)"
            } else {
                switch -Wildcard ($CalLog.TriggerAction) {
                    "Update" { $Action = "Updated" }
                    "Create" { $Action = "Sent" }
                    "*Delete*" { $Action = "Deleted" }
                    default {
                        $Action = "Updated"
                    }
                }

                $Extra = ""
                if (IsExceptionLog $CalLog) {
                    $Extra = " to the meeting starting $($CalLog.StartTime)"
                } elseif ($CalLog.AppointmentRecurring) {
                    $Extra = " to the meeting series"
                }

                if ($IsOrganizer) {
                    [array] $Output = "[$($CalLog.Organizer)] $($Action) a $($MeetingRespType) meeting Response message$($Extra)."
                } else {
                    switch ($CalLog.Client) {
                        ResourceBookingAssistant {
                            [array] $Output = "ResourceBookingAssistant $($Action) a $($MeetingRespType) Meeting Response message$($Extra)."
                        }
                        Transport {
                            [array] $Output = "[$($CalLog.From)] $($Action) $($MeetingRespType) Meeting Response message$($Extra)."
                        }
                        default {
                            [array] $Output = "[$($CalLog.ResponsibleUser)] $($Action) [$($CalLog.Organizer)]'s $($MeetingRespType) Meeting Response with $($CalLog.Client)$($Extra)."
                        }
                    }
                }
            }
        }
        Forward.Notification {
            [array] $Output = "The meeting was FORWARDED by [$($CalLog.Organizer)]."
        }
        Exception* {
            if ($CalLog.ResponsibleUser -ne "Calendar Assistant") {
                [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction)d Exception starting $($CalLog.StartTime) to the meeting series with $($CalLog.Client)."
            }
        }
        Ipm.Appointment {
            switch ($CalLog.TriggerAction) {
                Create {
                    if ($IsOrganizer) {
                        if ($CalLog.Client -eq "Transport") {
                            [array] $Output = "Transport Created a new meeting."
                        } else {
                            [array] $Output = "[$($CalLog.ResponsibleUser)] Created a new Meeting with $($CalLog.Client)."
                        }
                    } else {
                        switch ($CalLog.Client) {
                            Transport {
                                [array] $Output = "Transport Created a new Meeting on the calendar from [$($CalLog.Organizer)] and marked it Tentative."
                            }
                            ResourceBookingAssistant {
                                [array] $Output = "ResourceBookingAssistant Created a new Meeting on the calendar from [$($CalLog.Organizer)] and marked it Tentative."
                            }
                            default {
                                [array] $Output = "[$($CalLog.ResponsibleUser)] Created the Meeting with $($CalLog.Client)."
                            }
                        }
                    }
                }
                Update {
                    switch ($CalLog.Client) {
                        Transport {
                            if ($CalLog.ResponsibleUser -eq "Calendar Assistant") {
                                [array] $Output = "Transport Updated the meeting based on changes made to the meeting on [$($CalLog.Sender)] calendar."
                            } else {
                                [array] $Output = "Transport $($CalLog.TriggerAction)d the meeting based on changes made by [$($CalLog.ResponsibleUser)]."
                            }
                        }
                        LocationProcessor {
                            [array] $Output = ""
                        }
                        ResourceBookingAssistant {
                            [array] $Output = "ResourceBookingAssistant $($CalLog.TriggerAction)d the Meeting."
                        }
                        CalendarRepairAssistant {
                            [array] $Output = "CalendarRepairAssistant $($CalLog.TriggerAction)d the Meeting to repair an inconsistency."
                        }
                        default {
                            if ($CalLog.ResponsibleUser -eq "Calendar Assistant") {
                                [array] $Output = "The Exchange System $($CalLog.TriggerAction)d the meeting via the Calendar Assistant."
                            } else {
                                [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction)d the Meeting with $($CalLog.Client)."
                                $script:AddChangedProperties = $True
                            }
                        }
                    }

                    if ($CalLog.FreeBusyStatus -eq 2 -and $script:PreviousCalLog.FreeBusyStatus -ne 2) {
                        if ($CalLog.ResponsibleUserName -eq "Calendar Assistant") {
                            [array] $Output = "$($CalLog.Client) Accepted the meeting."
                        } else {
                            [array] $Output = "[$($CalLog.ResponsibleUser)] Accepted the meeting with $($CalLog.Client)."
                        }
                        $script:AddChangedProperties = $False
                    } elseif ($CalLog.FreeBusyStatus -ne 2 -and $script:PreviousCalLog.FreeBusyStatus -eq 2) {
                        if ($IsOrganizer) {
                            [array] $Output = "[$($CalLog.ResponsibleUser)] Cancelled the Meeting with $($CalLog.Client)."
                        } else {
                            if ($CalLog.ResponsibleUser -ne "Calendar Assistant") {
                                [array] $Output = "[$($CalLog.ResponsibleUser)] Declined the meeting with $($CalLog.Client)."
                            }
                        }
                        $script:AddChangedProperties = $False
                    }
                }
                SoftDelete {
                    switch ($CalLog.Client) {
                        Transport {
                            [array] $Output = "Transport $($CalLog.TriggerAction)d the meeting based on changes by [$($CalLog.Organizer)]."
                        }
                        LocationProcessor {
                            [array] $Output = ""
                        }
                        ResourceBookingAssistant {
                            [array] $Output = "ResourceBookingAssistant $($CalLog.TriggerAction)d the Meeting."
                        }
                        default {
                            if ($CalLog.ResponsibleUser -eq "Calendar Assistant") {
                                [array] $Output = "The Exchange System $($CalLog.TriggerAction)d the meeting via the Calendar Assistant."
                            } else {
                                [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction)d the meeting with $($CalLog.Client)."
                                $script:AddChangedProperties = $True
                            }
                        }
                    }

                    if ($CalLog.FreeBusyStatus -eq 2 -and $script:PreviousCalLog.FreeBusyStatus -ne 2) {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] Accepted the Meeting with $($CalLog.Client)."
                        $script:AddChangedProperties = $False
                    } elseif ($CalLog.FreeBusyStatus -ne 2 -and $script:PreviousCalLog.FreeBusyStatus -eq 2) {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] Declined the Meeting with $($CalLog.Client)."
                        $script:AddChangedProperties = $False
                    }
                }
                MoveToDeletedItems {
                    [array] $Output = "[$($CalLog.ResponsibleUser)] Deleted the Meeting with $($CalLog.Client) (Moved the Meeting to the Deleted Items)."
                }
                default {
                    [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction) the Meeting with $($CalLog.Client)."
                    $script:MeetingSummaryNeeded = $False
                }
            }
        }
        Cancellation {
            switch ($CalLog.Client) {
                Transport {
                    if ($CalLog.IsException -eq $True) {
                        [array] $Output = "Transport $($CalLog.TriggerAction)d a Meeting Cancellation based on changes by [$($CalLog.Sender)] for the exception starting on [$($CalLog.StartTime)]"
                    } else {
                        [array] $Output = "Transport $($CalLog.TriggerAction)d a Meeting Cancellation based on changes by [$($CalLog.Sender)]."
                    }
                }
                ResourceBookingAssistant {
                    if ($CalLog.TriggerAction -eq "MoveToDeletedItems") {
                        [array] $Output = "ResourceBookingAssistant Deleted the Cancellation."
                    } else {
                        [array] $Output = "ResourceBookingAssistant $($CalLog.TriggerAction)d the Cancellation."
                    }
                }
                default {
                    if ($CalLog.IsException -eq $True) {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction)d a Cancellation with $($CalLog.Client) for the exception starting on [$($CalLog.StartTime)]."
                    } elseif ($CalLog.CalendarItemType -eq "RecurringMaster") {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction)d a Cancellation for the Series with $($CalLog.Client)."
                    } else {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction)d the Cancellation with $($CalLog.Client)."
                    }
                }
            }
        }
        default {
            if ($CalLog.TriggerAction -eq "Create") {
                $Action = "New "
            } else {
                $Action = "$($CalLog.TriggerAction)"
            }
            [array] $Output = "[$($CalLog.ResponsibleUser)] performed a $($Action) on the $($CalLog.ItemClass) with $($CalLog.Client)."
        }
    }

    return $Output
}

<#
.SYNOPSIS
    Returns true if the value is effectively "not found" / empty.
#>
function IsNotFound {
    param($Value)
    return ([string]::IsNullOrEmpty($Value) -or $Value -eq "NotFound" -or $Value -eq "-")
}

<#
.SYNOPSIS
    Checks if a property changed between two CalLog entries, ignoring transitions to/from NotFound.
    Returns true if there is a meaningful change worth reporting.
#>
function HasMeaningfulChange {
    param(
        $OldValue,
        $NewValue
    )
    if ([string]::Equals([string]$OldValue, [string]$NewValue, [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
    if ((IsNotFound $OldValue) -or (IsNotFound $NewValue)) { return $false }
    return $true
}

<#
.SYNOPSIS
    Determines if key properties of the calendar log have changed.
.DESCRIPTION
    This function checks if the properties of the calendar log have changed by comparing the current
    Calendar log to the Previous calendar log (where it was an IPM.Appointment - i.e. the meeting)

    Changed properties will be added to the Timeline.
#>
function FindChangedProperties {
    if ($CalLog.Client -ne "LocationProcessor" -and $CalLog.Client -notlike "*EBA*" -and $CalLog.Client -notlike "*TBA*") {
        if ($script:PreviousCalLog -and $script:AddChangedProperties) {
            if (HasMeaningfulChange $script:PreviousCalLog.StartTime.ToString() $CalLog.StartTime.ToString()) {
                [Array]$TimeLineText = "The StartTime changed from [$($script:PreviousCalLog.StartTime)] to: [$($CalLog.StartTime)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.EndTime.ToString() $CalLog.EndTime.ToString()) {
                [Array]$TimeLineText = "The EndTime changed from [$($script:PreviousCalLog.EndTime)] to: [$($CalLog.EndTime)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.SubjectProperty $CalLog.SubjectProperty) {
                [Array]$TimeLineText = "The SubjectProperty changed from [$($script:PreviousCalLog.SubjectProperty)] to: [$($CalLog.SubjectProperty)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.NormalizedSubject $CalLog.NormalizedSubject) {
                [Array]$TimeLineText = "The NormalizedSubject changed from [$($script:PreviousCalLog.NormalizedSubject)] to: [$($CalLog.NormalizedSubject)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.Location $CalLog.Location) {
                [Array]$TimeLineText = "The Location changed from [$($script:PreviousCalLog.Location)] to: [$($CalLog.Location)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.TimeZone $CalLog.TimeZone) {
                [Array]$TimeLineText = "The TimeZone changed from [$($script:PreviousCalLog.TimeZone)] to: [$($CalLog.TimeZone)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.DisplayAttendeesAll $CalLog.DisplayAttendeesAll) {
                [Array]$TimeLineText = "The All Attendees changed from [$($script:PreviousCalLog.DisplayAttendeesAll)] to: [$($CalLog.DisplayAttendeesAll)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.AppointmentRecurring $CalLog.AppointmentRecurring) {
                [Array]$TimeLineText = "The Appointment Recurrence changed from [$($script:PreviousCalLog.AppointmentRecurring)] to: [$($CalLog.AppointmentRecurring)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.HasAttachment $CalLog.HasAttachment) {
                [Array]$TimeLineText = "The Meeting has Attachment changed from [$($script:PreviousCalLog.HasAttachment)] to: [$($CalLog.HasAttachment)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.IsCancelled $CalLog.IsCancelled) {
                [Array]$TimeLineText = "The Meeting is Cancelled changed from [$($script:PreviousCalLog.IsCancelled)] to: [$($CalLog.IsCancelled)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.IsAllDayEvent $CalLog.IsAllDayEvent) {
                [Array]$TimeLineText = "The Meeting is an All Day Event changed from [$($script:PreviousCalLog.IsAllDayEvent)] to: [$($CalLog.IsAllDayEvent)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.IsException $CalLog.IsException) {
                [Array]$TimeLineText = "The Meeting Is Exception changed from [$($script:PreviousCalLog.IsException)] to: [$($CalLog.IsException)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.IsSeriesCancelled $CalLog.IsSeriesCancelled) {
                [Array]$TimeLineText = "The Is Series Cancelled changed from [$($script:PreviousCalLog.IsSeriesCancelled)] to: [$($CalLog.IsSeriesCancelled)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.EventEmailReminderTimer $CalLog.EventEmailReminderTimer) {
                [Array]$TimeLineText = "The Email Reminder changed from [$($script:PreviousCalLog.EventEmailReminderTimer)] to: [$($CalLog.EventEmailReminderTimer)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.FreeBusyStatus $CalLog.FreeBusyStatus) {
                [Array]$TimeLineText = "The FreeBusy Status changed from [$($script:PreviousCalLog.FreeBusyStatus)] to: [$($CalLog.FreeBusyStatus)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.AppointmentState $CalLog.AppointmentState) {
                [Array]$TimeLineText = "The Appointment State changed from [$($script:PreviousCalLog.AppointmentState)] to: [$($CalLog.AppointmentState)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.MeetingRequestType $CalLog.MeetingRequestType) {
                [Array]$TimeLineText = "The Meeting Request Type changed from [$($script:PreviousCalLog.MeetingRequestType.Value)] to: [$($CalLog.MeetingRequestType.Value)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.CalendarItemType $CalLog.CalendarItemType) {
                [Array]$TimeLineText = "The Calendar Item Type changed from [$($script:PreviousCalLog.CalendarItemType)] to: [$($CalLog.CalendarItemType)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.ResponseType $CalLog.ResponseType) {
                [Array]$TimeLineText = "The ResponseType changed from [$($script:PreviousCalLog.ResponseType)] to: [$($CalLog.ResponseType)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if (HasMeaningfulChange $script:PreviousCalLog.Sender $CalLog.Sender) {
                if (-not (IsSameOrganizerIdentity $script:PreviousCalLog.Sender $CalLog.Sender)) {
                    [Array]$TimeLineText = "The Sender Email Address changed from [$($script:PreviousCalLog.Sender)] to: [$($CalLog.Sender)]"
                    CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
                }
            }

            if (HasMeaningfulChange $script:PreviousCalLog.From $CalLog.From) {
                if (-not (IsSameOrganizerIdentity $script:PreviousCalLog.From $CalLog.From)) {
                    [Array]$TimeLineText = "The From changed from [$($script:PreviousCalLog.From)] to: [$($CalLog.From)]"
                    CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
                }
            }

            if (HasMeaningfulChange $script:PreviousCalLog.ReceivedRepresenting $CalLog.ReceivedRepresenting) {
                [Array]$TimeLineText = "The Received Representing changed from [$($script:PreviousCalLog.ReceivedRepresenting)] to: [$($CalLog.ReceivedRepresenting)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }
        }
    }
}

<#
.SYNOPSIS
    Function to write a line of text surrounded by a dash line box.

.DESCRIPTION
    The Write-DashLineBoxColor function is used to create a quick and easy display around a line of text. It generates a box made of dash characters ("-") and displays the provided line of text inside the box.

.PARAMETER Line
    Specifies the line of text to be displayed inside the dash line box.

.PARAMETER Color
    Specifies the color of the dash line box and the text. The default value is "White".

.PARAMETER DashChar
    Specifies the character used to create the dash line. The default value is "-".

.EXAMPLE
    Write-DashLineBoxColor -Line "Hello, World!" -Color "Yellow" -DashChar "="
    Displays:
    ==============
    Hello, World!
    ==============
#>
function Write-DashLineBoxColor {
    [CmdletBinding()]
    param(
        [string[]]$Line,
        [string] $Color = "White",
        [char] $DashChar = "-"
    )
    $highLineLength = 0
    $Line | ForEach-Object { if ($_.Length -gt $highLineLength) { $highLineLength = $_.Length } }
    $dashLine = [string]::Empty
    1..$highLineLength | ForEach-Object { $dashLine += $DashChar }
    Write-Host
    Write-Host -ForegroundColor $Color $dashLine
    $Line | ForEach-Object { Write-Host -ForegroundColor $Color $_ }
    Write-Host -ForegroundColor $Color $dashLine
    Write-Host
}

# Default to Excel unless specified otherwise.
if (!$ExportToCSV.IsPresent) {
    Write-Host -ForegroundColor Yellow "Exporting to Excel."
    $script:ExportToExcel = $true

function Confirm-Administrator {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

    return $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )
}
    $script:IsAdministrator = Confirm-Administrator

# ===================================================================================================
# ImportExcel Functions
# see https://github.com/dfinke/ImportExcel for information on the module.
# ===================================================================================================
function CheckExcelModuleInstalled {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param ()

    if (Get-Command -Module ImportExcel) {
        Write-Host "ImportExcel module is already installed."
    } else {
        # This is slow, to the tune of ~10 seconds, but much more complete.
        # Check if ImportExcel module is installed
        $moduleInstalled = Get-Module -ListAvailable | Where-Object { $_.Name -eq 'ImportExcel' }

        if ($moduleInstalled) {
            Write-Host "ImportExcel module is already installed."
        } else {
            # Check if running with administrator rights
            if (-not $script:IsAdministrator) {
                Write-Host "Please run the script as an administrator to install the ImportExcel module."
                exit
            }

            # Ask user if they want to install the module
            if ($PSCmdlet.ShouldProcess('ImportExcel module', 'Import')) {
                Write-Verbose "Installing ImportExcel module..."

                # Install ImportExcel module
                Install-Module -Name ImportExcel -Force -AllowClobber

                Write-Host -ForegroundColor Green "Done. ImportExcel module is now installed."
                Write-Host -ForegroundColor Yellow "Please rerun the Script to get your Calendar Logs."
                exit
            }
        }
    }
}

# Export to Excel
function Export-CalLogExcel {
    Write-Host -ForegroundColor Cyan "Exporting Enhanced CalLogs to Excel Tab [$ShortId]..."
    $script:lastRow = 1000 # Default last row, will be updated later
    $script:firstRow = 3 # Row 1 is the Title, Row 2 is the Header
    $script:lastColumn = "AN" # Column AN is the last column in the Excel sheet

    $ExcelParamsArray = GetExcelParams -path $FileName -tabName $ShortId

    # Suppress EPPlus warnings and errors that occur when writing to an existing file
    # (AutoNameRange validation, named range conflicts, title character warnings)
    $savedErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'SilentlyContinue'
        $excel = $script:EnhancedCalLogs | Export-Excel @ExcelParamsArray -PassThru 3>$null
    } finally {
        $ErrorActionPreference = $savedErrorActionPreference
    }

    FormatHeader ($excel)
    CheckRows ($excel)

    # Set tab color to match the role (Organizer=Orange, Room=Green, Attendee=Blue)
    $excel.Workbook.Worksheets[$ShortId].TabColor = $script:TabColor

    Export-Excel -ExcelPackage $excel -WorksheetName $ShortId -MoveToStart

    # Log script info (will be positioned in the middle)
    LogScriptInfo

    # Export Raw Logs for Developer Analysis in collected order; do not apply the Enhanced tab sort here.
    Write-Host -ForegroundColor Cyan "Exporting Raw CalLogs to Excel Tab [$($ShortId + "_Raw")]..."
    try {
        $ErrorActionPreference = 'SilentlyContinue'
        $rawExcel = $script:GCDO | Export-Excel -Path $FileName -WorksheetName $($ShortId + "_Raw") -AutoFilter -FreezeTopRow -BoldTopRow -MoveToEnd -ClearSheet -PassThru 3>$null
    } finally {
        $ErrorActionPreference = $savedErrorActionPreference
    }
    $rawExcel.Workbook.Worksheets[$ShortId + "_Raw"].TabColor = $script:TabColor
    Export-Excel -ExcelPackage $rawExcel -WorksheetName $($ShortId + "_Raw")
}

function Format-ElapsedTime {
    param(
        [TimeSpan]$Elapsed
    )

    if ($Elapsed.TotalHours -ge 1) {
        return ('{0:00}:{1:00}:{2:00}' -f [int]$Elapsed.TotalHours, $Elapsed.Minutes, $Elapsed.Seconds)
    }

    return ('{0:00}:{1:00}' -f [int]$Elapsed.TotalMinutes, $Elapsed.Seconds)
}

function Get-ExceptionSummaryText {
    $exceptionLogs = @($script:GCDO | Where-Object {
            ($null -ne $_.OriginalStartDate -and $_.OriginalStartDate -ne 'NotFound' -and -not [string]::IsNullOrEmpty([string]$_.OriginalStartDate)) -or
            (IsExceptionItemType (GetItemType $_.ItemClass))
        })
    $exceptionCount = $exceptionLogs.Count

    switch ($script:ExceptionCollectionStatus) {
        'CollectedFast' {
            if ($AllExceptions.IsPresent) {
                return "Yes - $exceptionCount exception log(s) collected with FastExceptions for all Exception dates."
            }
            return "Yes - $exceptionCount exception log(s) collected with FastExceptions for the last 6 months by default."
        }
        'CollectedLegacyFallback' {
            if ($AllExceptions.IsPresent) {
                return "Yes - $exceptionCount exception log(s) collected after FastExceptions fell back to the legacy collector for all Exception dates."
            }
            return "Yes - $exceptionCount exception log(s) collected after FastExceptions fell back to the legacy collector for the last 6 months by default."
        }
        'Collected' {
            return "Yes - $exceptionCount exception log(s) collected."
        }
        'NoExceptionDates' {
            return "No - the AppointmentRecurrenceBlob did not expose matching Exception dates."
        }
        'NotRecurring' {
            return "No - the meeting is not recurring."
        }
        'SkippedBySwitch' {
            return "No - skipped because -NoExceptions was used."
        }
        'SkippedUntilMeetingIdChosen' {
            return "No - skipped because -ExceptionDate requires rerunning with -MeetingID."
        }
        'SkippedMultipleMeetingIds' {
            return "No - skipped because Subject search resolved to multiple MeetingIDs."
        }
        'NoMeetingId' {
            return "No - no MeetingID was resolved from the Subject search."
        }
        default {
            if (-not $Exceptions.IsPresent) {
                return "No - Exception collection was not requested."
            }
            if ($exceptionCount -gt 0) {
                return "Yes - $exceptionCount exception log(s) present in the collected output."
            }
            return "No - no exception logs were found in the collected output."
        }
    }
}

function Get-CollectionMethodDescription {
    if ($script:SubjectSearch) {
        if ($script:SubjectMeetingIdCount -eq 1) {
            if ($script:ExceptionCollectionStatus -eq 'CollectedFast') {
                if ($AllExceptions.IsPresent) {
                    return 'Subject search resolved to one MeetingID; Exception data collected with FastExceptions for all Exception dates. Tracking Logs still require -MeetingID.'
                }
                return 'Subject search resolved to one MeetingID; Exception data collected with FastExceptions for the last 6 months by default. Tracking Logs still require -MeetingID.'
            }
            if ($script:ExceptionCollectionStatus -eq 'CollectedLegacyFallback') {
                if ($AllExceptions.IsPresent) {
                    return 'Subject search resolved to one MeetingID; FastExceptions failed and the script fell back to the legacy Exception collector for all Exception dates. Tracking Logs still require -MeetingID.'
                }
                return 'Subject search resolved to one MeetingID; FastExceptions failed and the script fell back to the legacy Exception collector for the last 6 months by default. Tracking Logs still require -MeetingID.'
            }
            if ($script:ExceptionCollectionStatus -eq 'Collected') {
                return 'Subject search resolved to one MeetingID; Exception data collected by default. Tracking Logs still require -MeetingID.'
            }
            if ($script:ExceptionCollectionStatus -eq 'SkippedBySwitch') {
                return 'Subject search resolved to one MeetingID; Exception collection was skipped because -NoExceptions was used. Tracking Logs still require -MeetingID.'
            }
            if ($script:ExceptionCollectionStatus -eq 'SkippedUntilMeetingIdChosen') {
                return 'Subject search resolved to one MeetingID; targeted Exception collection was skipped because -ExceptionDate requires rerunning with -MeetingID. Tracking Logs still require -MeetingID.'
            }
            if ($script:ExceptionCollectionStatus -eq 'NoExceptionDates') {
                return 'Subject search resolved to one MeetingID; the AppointmentRecurrenceBlob did not expose any matching Exception dates to collect. Tracking Logs still require -MeetingID.'
            }
            if ($script:ExceptionCollectionStatus -eq 'NotRecurring') {
                return 'Subject search resolved to one MeetingID; Exception lookup ran, but the meeting is not recurring. Tracking Logs still require -MeetingID.'
            }
            return 'Subject search resolved to one MeetingID. Tracking Logs still require -MeetingID for full collection.'
        }
        if ($script:SubjectMeetingIdCount -gt 1) {
            return 'Subject search resolved to multiple MeetingIDs; Exception collection was skipped. Re-run with a specific -MeetingID to collect meeting exceptions.'
        }
        return 'Subject search did not resolve to a MeetingID.'
    }

    if (-not [string]::IsNullOrEmpty($ExceptionDate)) {
        return 'MeetingID search with targeted -ExceptionDate collection.'
    }
    if ($NoExceptions.IsPresent) {
        return 'MeetingID search with Exception collection disabled by -NoExceptions.'
    }
    if ($script:ExceptionCollectionStatus -eq 'CollectedFast') {
        if ($AllExceptions.IsPresent) {
            return 'MeetingID search with FastExceptions collection for all Exception dates.'
        }
        return 'MeetingID search with FastExceptions collection for the last 6 months by default.'
    }
    if ($script:ExceptionCollectionStatus -eq 'CollectedLegacyFallback') {
        if ($AllExceptions.IsPresent) {
            return 'MeetingID search where FastExceptions failed and the script fell back to the legacy Exception collector for all Exception dates.'
        }
        return 'MeetingID search where FastExceptions failed and the script fell back to the legacy Exception collector for the last 6 months by default.'
    }
    if ($script:ExceptionCollectionStatus -eq 'Collected') {
        return 'MeetingID search with legacy Exception collection.'
    }
    if ($script:ExceptionCollectionStatus -eq 'NoExceptionDates') {
        return 'MeetingID search where the AppointmentRecurrenceBlob did not expose any matching Exception dates to collect.'
    }
    if ($script:ExceptionCollectionStatus -eq 'NotRecurring') {
        return 'MeetingID search where the meeting is not recurring, so no Exception logs were collected.'
    }
    return 'MeetingID'
}

function LogScriptInfo {
    # Increment run number only once per script invocation (not per identity)
    if (-not $script:RunNumberSet) {
        # Read the existing run count from the Excel file so numbering is continuous across sessions
        if ($script:RunNumber -eq 0 -and (Test-Path $FileName)) {
            try {
                $pkg = Open-ExcelPackage -Path $FileName -ErrorAction SilentlyContinue
                if ($null -ne $pkg -and $null -ne $pkg.Workbook.Worksheets["Script Info"]) {
                    $existingInfo = Import-Excel -ExcelPackage $pkg -WorksheetName "Script Info" -ErrorAction SilentlyContinue
                    if ($null -ne $existingInfo) {
                        $lastRunKey = $existingInfo | Where-Object { $_.Key -like "--- Run #* ---" } | Select-Object -Last 1
                        if ($null -ne $lastRunKey -and $lastRunKey.Key -match '#(\d+)') {
                            $script:RunNumber = [int]$Matches[1]
                        }
                    }
                }
            } catch {
                Write-Verbose "Unable to read existing run count from Excel: $_"
            } finally {
                if ($null -ne $pkg) { $pkg.Dispose() }
            }
        }
        $script:RunNumber++
        $script:RunNumberSet = $true
        $script:RunHeaderWritten = $false
    }

    $RunInfo = [System.Collections.Generic.List[object]]::new()

    # Write run header and shared info only once per run (on the first identity)
    if (-not $script:RunHeaderWritten) {
        $RunInfo.Add([PSCustomObject]@{
                Key   = "--- Run #$($script:RunNumber) ---"
                Value = "---"
            })
        $RunInfo.Add([PSCustomObject]@{
                Key   = "RunTime"
                Value = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            })
        $RunInfo.Add([PSCustomObject]@{
                Key   = "Script Name"
                Value = $($script:command.MyCommand.Name)
            })
        $RunInfo.Add([PSCustomObject]@{
                Key   = "Command Line"
                Value = $($script:command.Line)
            })

        # Combine version and age into one line
        $versionText = if (-not [string]::IsNullOrEmpty($script:BuildVersion)) { $script:BuildVersion } else { "Unknown" }
        if (-not [string]::IsNullOrEmpty($script:BuildVersion)) {
            try {
                $parts = $script:BuildVersion.Split('.')
                $buildDate = [DateTime]::new(2000 + [int]$parts[0], [int]$parts[1], [int]$parts[2])
                $age = (Get-Date) - $buildDate
                $ageText = if ($age.Days -eq 0) { "today" }
                elseif ($age.Days -eq 1) { "1 day old" }
                elseif ($age.Days -lt 30) { "$($age.Days) days old" }
                elseif ($age.Days -lt 365) { "$([math]::Floor($age.Days / 30)) month(s) old" }
                else { "$([math]::Floor($age.Days / 365)) year(s), $([math]::Floor(($age.Days % 365) / 30)) month(s) old" }
                $versionText = "$versionText ($ageText, built $($buildDate.ToString('yyyy-MM-dd')))"
            } catch {
                Write-Verbose "Unable to parse BuildVersion '$($script:BuildVersion)' for age calculation: $_"
            }
        }
        $RunInfo.Add([PSCustomObject]@{
                Key   = "Script Version"
                Value = $versionText
            })

        $RunInfo.Add([PSCustomObject]@{
                Key   = "Identities"
                Value = ($ValidatedIdentities -join '; ')
            })

        $RunInfo.Add([PSCustomObject]@{
                Key   = "Collection Method"
                Value = (Get-CollectionMethodDescription)
            })

        # Only log environment info on the first run
        if ($script:RunNumber -eq 1) {
            $RunInfo.Add([PSCustomObject]@{
                    Key   = "User"
                    Value = whoami.exe
                })
            $RunInfo.Add([PSCustomObject]@{
                    Key   = "PowerShell Version"
                    Value = $PSVersionTable.PSVersion
                })
            $RunInfo.Add([PSCustomObject]@{
                    Key   = "OS Version"
                    Value = $(Get-CimInstance -ClassName Win32_OperatingSystem).Version
                })
            $RunInfo.Add([PSCustomObject]@{
                    Key   = "More Info"
                    Value = "https://learn.microsoft.com/en-us/exchange/troubleshoot/calendars/analyze-calendar-diagnostic-logs"
                })
        }

        $script:RunHeaderWritten = $true
    }

    # Per-identity line: show which identity was collected and log count
    $logCount = if ($null -ne $script:GCDO) { $script:GCDO.Count } else { 0 }
    $elapsedText = 'Unknown'
    if ($null -ne $script:CurrentIdentityRunStartTime) {
        $elapsedText = Format-ElapsedTime -Elapsed ((Get-Date) - $script:CurrentIdentityRunStartTime)
    }
    $RunInfo.Add([PSCustomObject]@{
            Key   = "  $($script:Identity)"
            Value = "$logCount logs collected in $elapsedText at $(Get-Date -Format 'HH:mm:ss')"
        })
    $RunInfo.Add([PSCustomObject]@{
            Key   = "    Duration"
            Value = $elapsedText
        })
    $RunInfo.Add([PSCustomObject]@{
            Key   = "    Exceptions"
            Value = (Get-ExceptionSummaryText)
        })

    # Capture errors that occurred during this identity (new errors since last snapshot)
    # Filter out EPPlus internal errors (IsValidAddress) and Import-Excel worksheet-not-found errors
    $newErrors = @()
    for ($i = 0; $i -lt ($Error.Count - $script:PreRunErrorCount); $i++) {
        $err = $Error[$i]
        $msg = if ($err -is [System.Management.Automation.ErrorRecord]) { $err.Exception.Message } else { $err.ToString() }
        if ($msg -notlike '*IsValidAddress*' -and $msg -notlike '*Worksheet*Script Info*not found*') {
            $newErrors += $err
        }
    }
    if ($newErrors.Count -gt 0) {
        for ($i = 0; $i -lt $newErrors.Count; $i++) {
            $err = $newErrors[$i]
            $errorMessage = if ($err -is [System.Management.Automation.ErrorRecord]) {
                "$($err.Exception.Message) [$($err.CategoryInfo.Category): $($err.FullyQualifiedErrorId)]"
            } else {
                $err.ToString()
            }
            $RunInfo.Add([PSCustomObject]@{
                    Key   = "    Error $($i + 1)"
                    Value = $errorMessage
                })
        }
    }
    # Update snapshot so next identity only captures new errors
    $script:PreRunErrorCount = $Error.Count

    # Write Script Info into the existing workbook (Enhanced tab was already saved)
    $savedEAP = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'SilentlyContinue'

        # Open existing package and add/append the Script Info sheet
        try {
            $pkg = Open-ExcelPackage -Path $FileName -ErrorAction Stop
            $sheetExists = $null -ne $pkg.Workbook.Worksheets["Script Info"]

            if ($sheetExists) {
                # Append rows after existing data
                $ws = $pkg.Workbook.Worksheets["Script Info"]
                $startRow = $ws.Dimension.End.Row + 1
            } else {
                # Create a new sheet
                $ws = $pkg.Workbook.Worksheets.Add("Script Info")
                $startRow = 1
                # Add header row
                $ws.Cells[$startRow, 1].Value = "Key"
                $ws.Cells[$startRow, 2].Value = "Value"
                $startRow = 2
            }

            foreach ($item in $RunInfo) {
                $ws.Cells[$startRow, 1].Value = $item.Key
                $ws.Cells[$startRow, 2].Value = [string]$item.Value
                $startRow++
            }

            $ws.Column(1).Width = 25
            $ws.Column(2).Width = 80
            $ws.TabColor = [System.Drawing.Color]::Gray
            $pkg.Save()
            $pkg.Dispose()
        } catch {
            Write-Warning "Unable to write Script Info tab: $_"
            if ($null -ne $pkg) { $pkg.Dispose() }
        }
    } finally {
        $ErrorActionPreference = $savedEAP
    }
}

function Export-TimelineExcel {
    Write-Host -ForegroundColor Cyan "Exporting Timeline to Excel..."
    $savedEAP = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'SilentlyContinue'
        $tlExcel = $script:TimeLineOutput | Export-Excel -Path $FileName -WorksheetName $($ShortId + "_TimeLine") -Title "Timeline for $Identity" -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet -PassThru 3>$null
    } finally {
        $ErrorActionPreference = $savedEAP
    }
    $tlExcel.Workbook.Worksheets[$ShortId + "_TimeLine"].TabColor = $script:TabColor
    Export-Excel -ExcelPackage $tlExcel -WorksheetName $($ShortId + "_TimeLine")
}

function GetExcelParams($path, $tabName) {
    if ($script:IsOrganizer) {
        $TableStyle = "Light10" # Orange for Organizer
        $TitleExtra = ", Organizer"
        $script:TabColor = [System.Drawing.Color]::FromArgb(237, 125, 49) # Orange
    } elseif ($script:IsRoomMB) {
        Write-Host -ForegroundColor green "Room Mailbox Detected"
        $TableStyle = "Light11" # Green for Room Mailbox
        $TitleExtra = ", Resource"
        $script:TabColor = [System.Drawing.Color]::FromArgb(112, 173, 71) # Green
    } else {
        $TableStyle = "Light12" # Light Blue for normal
        # Dark Blue for Delegates (once we can determine this)
        $script:TabColor = [System.Drawing.Color]::FromArgb(91, 155, 213) # Blue
    }

    if ($script:CalLogsDisabled) {
        $TitleExtra += ", WARNING: CalLogs are Turned Off for $Identity! This will be a incomplete story"
        $script:TabColor = [System.Drawing.Color]::FromArgb(255, 0, 0) # Red for Disabled CalLogs
    }

    if ($script:SubjectSearch) {
        if ($script:SubjectMeetingIdCount -eq 1) {
            if ($script:ExceptionCollectionStatus -eq "CollectedFast") {
                $TitleExtra += ", Collected with Subject search and FastExceptions"
            } elseif ($script:ExceptionCollectionStatus -eq "CollectedLegacyFallback") {
                $TitleExtra += ", Collected with Subject search (FastExceptions fell back to legacy)"
            } elseif ($script:ExceptionCollectionStatus -eq "Collected") {
                $TitleExtra += ", Collected with Subject search and Exception data"
            } elseif ($script:ExceptionCollectionStatus -eq "SkippedBySwitch") {
                $TitleExtra += ", Collected with Subject search (-NoExceptions used)"
            } elseif ($script:ExceptionCollectionStatus -eq "SkippedUntilMeetingIdChosen") {
                $TitleExtra += ", Collected with Subject search (-ExceptionDate requires rerun with -MeetingID)"
            } elseif ($script:ExceptionCollectionStatus -eq "NoExceptionDates") {
                $TitleExtra += ", Collected with Subject search (no matching Exception dates in AppointmentRecurrenceBlob)"
            } elseif ($script:ExceptionCollectionStatus -eq "NotRecurring") {
                $TitleExtra += ", Collected with Subject search (no recurring Exception data)"
            } else {
                $TitleExtra += ", Collected with Subject search"
            }
        } elseif ($script:SubjectMeetingIdCount -gt 1) {
            $TitleExtra += ", Collected with Subject search (multiple MeetingIDs; rerun with -MeetingID for Exceptions)"
        } else {
            $TitleExtra += ", Collected with Subject search"
        }
    }

    $script:lastRow = $script:GCDO.Count + $firstRow - 1 # Last row is the number of items in the GCDO array + 2 for the header and title rows.
    Write-Host -ForegroundColor Gray "Last Row is $lastRow, First Row is $firstRow, Last Column is $lastColumn"

    return @{
        Path                    = $path
        FreezeTopRow            = $true
        #  BoldTopRow              = $true
        Verbose                 = $false
        TableStyle              = $TableStyle
        WorksheetName           = $tabName
        TableName               = $tabName
        FreezeTopRowFirstColumn = $true
        AutoFilter              = $true
        ClearSheet              = $true
        Title                   = "Enhanced Calendar Logs for $Identity" + $TitleExtra + " for MeetingID [$($script:GCDO[0].CleanGlobalObjectId)]."
        TitleSize               = 14
        ConditionalText         = $ConditionalFormatting
    }
}

$ColumnMap = @{
    LogTimestamp                   = "A"
    LogRowType                     = "B"
    SubjectProperty                = "C"
    Client                         = "D"
    LogClientInfoString            = "E"
    TriggerAction                  = "F"
    ItemClass                      = "G"
    SeqExpItemVersion              = "H"
    Organizer                      = "I"
    From                           = "J"
    FreeBusyStatus                 = "K"
    ResponsibleUser                = "L"
    Sender                         = "M"
    LogFolder                      = "N"
    OriginalLogFolder              = "O"
    SharedFolderName               = "P"
    ReceivedRepresenting           = "Q"
    MeetingRequestType             = "R"
    StartTime                      = "S"
    EndTime                        = "T"
    OriginalStartDate              = "U"
    Location                       = "V"
    CalendarItemType               = "W"
    RecurrencePattern              = "X"
    AppointmentAuxiliaryFlags      = "Y"
    DisplayAttendeesAll            = "Z"
    AttendeeCount                  = "AA"
    AppointmentState               = "AB"
    ResponseType                   = "AC"
    ClientIntent                   = "AD"
    AppointmentRecurring           = "AE"
    HasAttachment                  = "AF"
    IsCancelled                    = "AG"
    IsAllDayEvent                  = "AH"
    Sensitivity                    = "AI"
    IsSeriesCancelled              = "AJ"
    SendMeetingMessagesDiagnostics = "AK"
    AttendeeCollection             = "AL"
    CalendarLogRequestId           = "AM"
    CleanGlobalObjectId            = "AN"
}

function GetExcelColumnNumber {
    param([string]$ColumnLetter)
    $number = 0
    $letters = $ColumnLetter.ToUpper().ToCharArray()
    foreach ($char in $letters) {
        $number = $number * 26 + ([int][char]$char - [int][char]'A' + 1)
    }
    return $number
}

function Get-ColumnRange {
    param(
        [string]$PropertyName,
        [int]$StartRow = 3,
        [int]$EndRow = $script:lastRow
    )

    if ($EndRow -eq 0) {
        $EndRow = 2000
    }

    $col = $ColumnMap[$PropertyName]

    # if ($null -eq $col) { throw "Unknown property: $PropertyName" }
    if ($StartRow -and $EndRow) {
        #  Write-Host -ForegroundColor DarkGray "Getting column range for $PropertyName : $col, StartRow: $StartRow, EndRow: $EndRow"
        return $col + $StartRow + ":" + $col + $EndRow
    } else {
        return $col + ":" + $col
    }
}

$ConditionalFormatting = @(
    # Client, ShortClientInfoString and LogClientInfoString
    New-ConditionalText "Outlook" -ConditionalTextColor Green -BackgroundColor $null
    New-ConditionalText "OWA" -ConditionalTextColor DarkGreen -BackgroundColor $null
    New-ConditionalText "Teams" -ConditionalTextColor DarkGreen -BackgroundColor $null
    New-ConditionalText "Transport" -ConditionalTextColor Blue -BackgroundColor $null
    New-ConditionalText "Repair" -ConditionalTextColor DarkRed -BackgroundColor LightPink
    New-ConditionalText "Other ?BA" -ConditionalTextColor Orange -BackgroundColor $null
    New-ConditionalText "TimeService" -ConditionalTextColor Orange -BackgroundColor $null
    New-ConditionalText "Other REST" -ConditionalTextColor DarkRed -BackgroundColor $null
    New-ConditionalText "Unknown" -ConditionalTextColor DarkRed -BackgroundColor $null
    New-ConditionalText "ResourceBookingAssistant" -ConditionalTextColor Blue -BackgroundColor $null
    New-ConditionalText "Calendar Replication" -ConditionalTextColor Blue -BackgroundColor $null

    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'LogRowType') -ConditionalType ContainsText -Text "Interesting" -ConditionalTextColor Green -BackgroundColor $null
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'LogRowType') -ConditionalType ContainsText -Text "SeriesException" -ConditionalTextColor Green -BackgroundColor $null
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'LogRowType') -ConditionalType ContainsText -Text "DeletedSeriesException" -ConditionalTextColor Orange -BackgroundColor $null
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'LogRowType') -ConditionalType ContainsText -Text "MeetingMessageChange" -ConditionalTextColor Orange -BackgroundColor $null
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'LogRowType') -ConditionalType ContainsText -Text "SyncOrReplication" -ConditionalTextColor Blue -BackgroundColor $null
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'LogRowType') -ConditionalType ContainsText -Text "OtherAssistant" -ConditionalTextColor Orange -BackgroundColor $null

    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'TriggerAction') -ConditionalType ContainsText -Text "Create" -ConditionalTextColor Green -BackgroundColor $null
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'TriggerAction') -ConditionalType ContainsText -Text "Delete" -ConditionalTextColor Red -BackgroundColor $null

    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'ItemClass') -ConditionalType ContainsText -Text "IPM.Appointment" -ConditionalTextColor Blue -BackgroundColor $null
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'ItemClass') -ConditionalType ContainsText -Text "Cancellation" -ConditionalTextColor Black -BackgroundColor Orange
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'ItemClass') -ConditionalType ContainsText -Text ".Request" -ConditionalTextColor DarkGreen -BackgroundColor $null
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'ItemClass') -ConditionalType ContainsText -Text ".Resp." -ConditionalTextColor Orange -BackgroundColor $null
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'ItemClass') -ConditionalType ContainsText -Text "IPM.OLE.CLASS" -ConditionalTextColor Plum -BackgroundColor $null

    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'FreeBusyStatus') -ConditionalType ContainsText -Text "Free" -ConditionalTextColor Red -BackgroundColor $null
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'FreeBusyStatus') -ConditionalType ContainsText -Text "Tentative" -ConditionalTextColor Orange -BackgroundColor $null
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'FreeBusyStatus') -ConditionalType ContainsText -Text "Busy" -ConditionalTextColor Green -BackgroundColor $null

    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'MeetingRequestType') -ConditionalType ContainsText -Text "Outdated" -ConditionalTextColor DarkRed -BackgroundColor LightPink

    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'CalendarItemType') -ConditionalType ContainsText -Text "RecurringMaster" -ConditionalTextColor $null -BackgroundColor Plum

    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'AppointmentAuxiliaryFlags') -ConditionalType ContainsText -Text "Copied" -ConditionalTextColor DarkRed -BackgroundColor LightPink
    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'AppointmentAuxiliaryFlags') -ConditionalType ContainsText -Text "ForwardedAppointment" -ConditionalTextColor DarkRed -BackgroundColor $null

    New-ConditionalText -Range (Get-ColumnRange -PropertyName 'ResponseType') -ConditionalType ContainsText -Text "Organizer" -ConditionalTextColor Orange -BackgroundColor $null
)

function CheckRows {
    param(
        [object] $excel
    )
    $sheet = $excel.Workbook.Worksheets[$ShortId]

    # Pre-filter LogRowType to hide noisy rows (OtherAssistant, MeetingMessageChange, deleted exceptions)
    SetLogRowTypeFilter -excel $excel

    # Highlight the Resp in LightGoldenRodYellow
    CheckColumnForText -sheet $sheet -columnNumber $(GetExcelColumnNumber($ColumnMap.ItemClass)) -textToFind "Resp" -cellColor "LightGoldenRodYellow" -fontColor "Black"

    # Highlight the RUM in Red
    CheckColumnForText -sheet $sheet -columnNumber $(GetExcelColumnNumber($ColumnMap.AppointmentAuxiliaryFlags)) -textToFind "RepairUpdateMessage" -cellColor "White" -fontColor "DarkRed"

    #Highlight the Cancellation in Orange
    CheckColumnsForValues -sheet $sheet -columnNumber1  $(GetExcelColumnNumber($ColumnMap.ItemClass)) -value1 "Cancellation" -columnNumber2  $(GetExcelColumnNumber($ColumnMap.TriggerAction)) -value2 "Create" -cellColor "Khaki" -fontColor "Black"

    # Highlight the Create from Transport in light blue
    CheckColumnsForValues -sheet $sheet -columnNumber1  $(GetExcelColumnNumber($ColumnMap.LogClientInfoString)) -value1 "Transport" -columnNumber2  $(GetExcelColumnNumber($ColumnMap.TriggerAction)) -value2 "Create" -cellColor "LightBlue" -fontColor "Black"

    # Highlight SharedFolderName with unique colors per shared folder
    HighlightSharedFolderNames -sheet $sheet

    # Detect organizer anomalies (only on organizer's tab)
    if ($script:IsOrganizer) {
        CheckOrganizerErrors -sheet $sheet
    }

    $excel.Save()
}

<#
.SYNOPSIS
Pre-filters the LogRowType column in the Excel table to hide noisy row types.
Hides OtherAssistant, MeetingMessageChange, and DeletedSeriesException rows by default.
Users can clear the filter in Excel to see all rows.
#>
function SetLogRowTypeFilter {
    param(
        [object] $excel
    )

    $sheet = $excel.Workbook.Worksheets[$ShortId]
    $logRowCol = GetExcelColumnNumber($ColumnMap.LogRowType)

    # Values to hide by default
    $hideValues = @("OtherAssistant", "MeetingMessageChange", "DeletedSeriesException")

    # Collect all unique values in the LogRowType column
    $lastDataRow = $sheet.Dimension.End.Row
    $allValues = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    for ($row = $script:firstRow; $row -le $lastDataRow; $row++) {
        $val = $sheet.Cells[$row, $logRowCol].Text
        if (-not [string]::IsNullOrEmpty($val)) {
            [void]$allValues.Add($val)
        }
    }

    $showValues = @($allValues | Where-Object { $_ -notin $hideValues })
    if ($showValues.Count -eq $allValues.Count) {
        Write-Verbose "No LogRowType values to filter out."
        return
    }

    Write-Verbose "Pre-filtering LogRowType: hiding $($hideValues -join ', ')"

    # Get the table and its autoFilter XML node
    $tbl = $sheet.Tables[0]
    if ($null -eq $tbl) {
        Write-Verbose "No table found, skipping LogRowType filter."
        return
    }

    $ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    $nsm = [System.Xml.XmlNamespaceManager]::new($tbl.TableXml.NameTable)
    $nsm.AddNamespace("t", $ns)
    $autoFilter = $tbl.TableXml.SelectSingleNode("//t:autoFilter", $nsm)

    if ($null -eq $autoFilter) {
        Write-Verbose "No autoFilter found in table, skipping."
        return
    }

    # Build filterColumn with visible values (Excel filters list what to SHOW)
    # colId is 0-based relative to the table, LogRowType is the 2nd column (B) = colId 1
    $tableStartCol = $tbl.Address.Start.Column
    $filterColId = $logRowCol - $tableStartCol

    # Remove any existing filterColumn for this colId to avoid duplicates on reruns
    $existingFilterCol = $autoFilter.SelectSingleNode("t:filterColumn[@colId='$filterColId']", $nsm)
    if ($null -ne $existingFilterCol) {
        $null = $autoFilter.RemoveChild($existingFilterCol)
    }

    $filterCol = $tbl.TableXml.CreateElement("filterColumn", $ns)
    $filterCol.SetAttribute("colId", $filterColId.ToString())
    $filters = $tbl.TableXml.CreateElement("filters", $ns)

    foreach ($val in $showValues) {
        $f = $tbl.TableXml.CreateElement("filter", $ns)
        $f.SetAttribute("val", $val)
        $null = $filters.AppendChild($f)
    }
    $null = $filterCol.AppendChild($filters)
    $null = $autoFilter.AppendChild($filterCol)

    # Hide the actual rows so they appear filtered on open
    for ($row = $script:firstRow; $row -le $lastDataRow; $row++) {
        $val = $sheet.Cells[$row, $logRowCol].Text
        if ($val -in $hideValues) {
            $sheet.Row($row).Hidden = $true
        }
    }
}

<#
.SYNOPSIS
Detects organizer anomalies on the Organizer's tab:
1. ResponseType changes from "Organizer" to something else on IPM.Appointment/Exception rows (not shared).
2. The From (Organizer SMTP) changes between qualifying rows, indicating multiple organizers.
Highlights the row text in Red and the problematic cell in Yellow.
#>
function CheckOrganizerErrors {
    param(
        [object] $sheet
    )

    $itemClassCol = GetExcelColumnNumber($ColumnMap.ItemClass)
    $sharedFolderCol = GetExcelColumnNumber($ColumnMap.SharedFolderName)
    $responseTypeCol = GetExcelColumnNumber($ColumnMap.ResponseType)
    $fromCol = GetExcelColumnNumber($ColumnMap.From)
    $lastDataRow = $sheet.Dimension.End.Row
    $lastDataCol = $sheet.Dimension.End.Column

    $firstOrganizer = $null  # Track the first From value seen on a qualifying row

    for ($row = $script:firstRow; $row -le $lastDataRow; $row++) {
        $itemClass = $sheet.Cells[$row, $itemClassCol].Text
        $sharedFolder = $sheet.Cells[$row, $sharedFolderCol].Text

        # Only check IPM.Appointment and Exception rows that are not shared
        if (($itemClass -ne "Ipm.Appointment" -and $itemClass -notlike "Exception*") -or
            $sharedFolder -ne "Not Shared") {
            continue
        }

        $responseType = $sheet.Cells[$row, $responseTypeCol].Text
        $fromValue = $sheet.Cells[$row, $fromCol].Text

        # Check 1: ResponseType should be "Organizer" on the organizer's tab
        if ($responseType -ne "Organizer") {
            # Red text for the whole row
            $rowRange = $sheet.Cells[$row, 1, $row, $lastDataCol]
            $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Red)
            # Yellow highlight on the ResponseType cell
            $cell = $sheet.Cells[$row, $responseTypeCol]
            $cell.Style.Fill.PatternType = 'Solid'
            $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
        }

        # Check 2: From (Organizer SMTP) should not change between qualifying rows
        if ([string]::IsNullOrEmpty($fromValue) -or $fromValue -eq "-") {
            continue
        }

        if ($null -eq $firstOrganizer) {
            $firstOrganizer = $fromValue
        } elseif (-not (IsSameOrganizerIdentity $fromValue $firstOrganizer)) {
            # Red text for the whole row
            $rowRange = $sheet.Cells[$row, 1, $row, $lastDataCol]
            $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::Red)
            # Yellow highlight on the From cell
            $cell = $sheet.Cells[$row, $fromCol]
            $cell.Style.Fill.PatternType = 'Solid'
            $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
        }
    }
}

<#
.SYNOPSIS
Highlights rows in the SharedFolderName column with a unique pale color per shared folder name.
Cycles through 7 pale colors if there are more than 7 unique shared folder names.
#>
function HighlightSharedFolderNames {
    param(
        [object] $sheet
    )

    $sharedFolderCol = GetExcelColumnNumber($ColumnMap.SharedFolderName)
    $lastDataRow = $sheet.Dimension.End.Row
    $lastDataCol = $sheet.Dimension.End.Column

    # Debug: verify column alignment
    $headerValue = $sheet.Cells[2, $sharedFolderCol].Text
    Write-Verbose "HighlightSharedFolderNames: Column $($ColumnMap.SharedFolderName) = col# $sharedFolderCol, header='$headerValue', rows=$($script:firstRow)..$lastDataRow"
    if ($headerValue -ne "SharedFolderName") {
        # Column map is misaligned — find the actual column by scanning the header row
        for ($c = 1; $c -le $lastDataCol; $c++) {
            if ($sheet.Cells[2, $c].Text -eq "SharedFolderName") {
                Write-Host -ForegroundColor Yellow "HighlightSharedFolderNames: SharedFolderName found at column $c (expected $sharedFolderCol)"
                $sharedFolderCol = $c
                break
            }
        }
    }

    # Pale colors that are readable with black text
    # cSpell:ignore FromArgb
    [System.Drawing.Color[]] $PaleColors = @(
        [System.Drawing.Color]::FromArgb(220, 245, 220)  # Pale Green
        [System.Drawing.Color]::FromArgb(180, 215, 255)  # Pale Blue
        [System.Drawing.Color]::FromArgb(255, 240, 220)  # Pale Peach
        [System.Drawing.Color]::FromArgb(245, 230, 255)  # Pale Lavender
        [System.Drawing.Color]::FromArgb(255, 255, 210)  # Pale Yellow
        [System.Drawing.Color]::FromArgb(255, 225, 230)  # Pale Pink
        [System.Drawing.Color]::FromArgb(220, 245, 245)  # Pale Cyan
    )

    # Collect unique SharedFolderName values and apply formatting in a single pass
    $folderNames = @{}
    $colorIndex = 0

    for ($row = $script:firstRow; $row -le $lastDataRow; $row++) {
        $cellValue = $sheet.Cells[$row, $sharedFolderCol].Text
        if ([string]::IsNullOrEmpty($cellValue) -or $cellValue -eq "Not Shared") {
            continue
        }
        if (-not $folderNames.ContainsKey($cellValue)) {
            $folderNames[$cellValue] = $PaleColors[$colorIndex % $PaleColors.Count]
            $colorIndex++
        }

        $color = $folderNames[$cellValue]
        # Italicize the entire row
        $rowRange = $sheet.Cells[$row, 1, $row, $lastDataCol]
        $rowRange.Style.Font.Italic = $true
        # Highlight only the SharedFolderName cell
        $cell = $sheet.Cells[$row, $sharedFolderCol]
        $cell.Style.Fill.PatternType = 'Solid'
        $cell.Style.Fill.BackgroundColor.SetColor($color)
    }

    if ($folderNames.Count -eq 0) {
        Write-Verbose "No shared folder names found to highlight."
    } else {
        Write-Verbose "Highlighted $($folderNames.Count) unique shared folder name(s)."
    }
}

function SortByLogTimestamp {
    param(
        [object] $excel
    )
    $sheet = $excel.Workbook.Worksheets[$ShortId]
    $dataStartRow = $script:firstRow # Row 3 (row 1 = title, row 2 = header)
    $dataEndRow = $sheet.Dimension.End.Row
    $dataEndCol = $sheet.Dimension.End.Column
    $sortCol = GetExcelColumnNumber($ColumnMap.LogTimestamp)

    # Sort the data range by LogTimestamp column ascending
    # EPPlus Sort() takes a 0-based column offset within the range
    $sortRange = $sheet.Cells[$dataStartRow, 1, $dataEndRow, $dataEndCol]
    $sortRange.Sort($sortCol - 1)
    $excel.Save()
}

# check if a column contains a specific text and highlight the row
# This function highlights a row in the Excel sheet based on the row number and specified colors.
# Parameters:
#   - $sheet: The Excel worksheet object.
function CheckColumnForText {
    param (
        [object] $sheet,
        [int] $columnNumber,
        [string] $textToFind,
        [string] $cellColor = "Yellow",
        [string] $fontColor = "DarkRed"
    )

    $lastDataRow = $sheet.Dimension.End.Row
    $lastDataCol = $sheet.Dimension.End.Column
    Write-Verbose "Checking column $columnNumber for text '$textToFind'..."
    for ($row = $script:firstRow; $row -le $lastDataRow; $row++) {
        $cellValue = $sheet.Cells[$row, $columnNumber].Text

        if ($cellValue -like "*$textToFind*") {
            HighlightRow -sheet $sheet -rowNumber $row -lastCol $lastDataCol -cellColor $cellColor -fontColor $fontColor
        }
    }
}

# Checks if two columns in the same row match specified values and highlights the row if both match.
function CheckColumnsForValues {
    param (
        [object] $sheet,
        [int] $columnNumber1,
        [string] $value1,
        [int] $columnNumber2,
        [string] $value2,
        [string] $cellColor = "LightPink",
        [string] $fontColor = "DarkRed"
    )

    $lastDataRow = $sheet.Dimension.End.Row
    $lastDataCol = $sheet.Dimension.End.Column
    Write-Verbose "Checking for rows where column $columnNumber1 = '$value1' AND column $columnNumber2 = '$value2'..."
    for ($row = $script:firstRow; $row -le $lastDataRow; $row++) {
        $cellValue1 = $sheet.Cells[$row, $columnNumber1].Text
        $cellValue2 = $sheet.Cells[$row, $columnNumber2].Text

        if ($cellValue1 -like "*$value1*" -and $cellValue2 -like "*$value2*") {
            HighlightRow -sheet $sheet -rowNumber $row -lastCol $lastDataCol -cellColor $cellColor -fontColor $fontColor
        }
    }
}

function HighlightRow {
    param(
        [object] $sheet,
        [int] $rowNumber,
        [int] $lastCol,
        [string] $cellColor = "Thistle",
        [string] $fontColor = "DarkRed"
    )

    $rowRange = $sheet.Cells[$rowNumber, 1, $rowNumber, $lastCol]
    Write-Verbose "Highlighting row $rowNumber with cell color [$cellColor] and font color [$fontColor]"
    $rowRange.Style.Fill.PatternType = 'Solid'
    $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::$cellColor)
    $rowRange.Style.Font.Color.SetColor([System.Drawing.Color]::$fontColor)
}

function FormatHeader {
    param(
        [object] $excel
    )
    $sheet = $excel.Workbook.Worksheets[$ShortId]
    $HeaderRow = 2

    # Define header metadata: property name, width, alignment, number format, comment
    $headerMeta = @(
        @{ Name = "LogTimestamp"; Width = 20; Align = "Center"; NumberFormat = "m/d/yyyy h:mm:ss"; Comment = "LogTimestamp: Time when the change was recorded in the CalLogs. This and all Times are in UTC." }
        @{ Name = "LogRowType"; Width = 20; Align = "Left"; Comment = "LogRowType: Filtered by default! Interesting logs are what to focus on. OtherAssistant, MeetingMessageChange, and DeletedSeriesException are hidden. Clear the filter to see all rows." }
        @{ Name = "SubjectProperty"; Width = 20; Align = "Left"; Comment = "SubjectProperty: The Subject of the Meeting." }
        @{ Name = "Client"; Width = 20; Align = "Left"; Comment = "Client (ShortClientInfoString): The 'friendly' Client name of the client that made the change." }
        @{ Name = "LogClientInfoString"; Width = 5; Align = "Left"; Comment = "LogClientInfoString: Full Client Info String of client that made the change." }
        @{ Name = "TriggerAction"; Width = 12; Align = "Center"; Comment = "TriggerAction (CalendarLogTriggerAction): The type of action that caused the change." }
        @{ Name = "ItemClass"; Width = 18; Align = "Left"; Comment = "ItemClass: The Class of the Calendar Item" }
        @{ Name = "SeqExpItemVersion"; Width = 10; Align = "Center"; Comment = "Seq:Exp:ItemVersion (AppointmentLastSequenceNumber:AppointmentSequenceNumber:ItemVersion): The Sequence Version, the Exception Version, and the Item Version.  Each type of item has its own count." }
        @{ Name = "Organizer"; Width = 20; Align = "Left"; Comment = "Organizer (From.FriendlyDisplayName): The Organizer of the Calendar Item." }
        @{ Name = "From"; Width = 20; Align = "Left"; Comment = "From: The SMTP address of the Organizer of the Calendar Item." }
        @{ Name = "FreeBusyStatus"; Width = 12; Align = "Center"; Comment = "FreeBusy (FreeBusyStatus): The FreeBusy Status of the Calendar Item." }
        @{ Name = "ResponsibleUser"; Width = 20; Align = "Left"; Comment = "ResponsibleUser(ResponsibleUserName): The Responsible User of the change." }
        @{ Name = "Sender"; Width = 20; Align = "Left"; Comment = "Sender (SenderEmailAddress): The Sender of the change." }
        @{ Name = "LogFolder"; Width = 16; Align = "Left"; Comment = "LogFolder (ParentDisplayName): The Log Folder that the CalLog was in." }
        @{ Name = "OriginalLogFolder"; Width = 16; Align = "Left"; Comment = "OriginalLogFolder (OriginalParentDisplayName): The Original Log Folder that the item was in / delivered to." }
        @{ Name = "SharedFolderName"; Width = 15; Align = "Left"; Comment = "SharedFolderName: Was this from a Modern Sharing, and if so what Folder." }
        @{ Name = "ReceivedRepresenting"; Width = 10; Align = "Left"; Comment = "ReceivedRepresenting: Who the item was Received for, often the Delegate." }
        @{ Name = "MeetingRequestType"; Width = 10; Align = "Center"; Comment = "MeetingRequestType: The Meeting Request Type of the Meeting." }
        @{ Name = "StartTime"; Width = 23; Align = "Center"; NumberFormat = "m/d/yyyy h:mm:ss"; Comment = "StartTime: The Start Time of the Meeting. This and all Times are in UTC." }
        @{ Name = "EndTime"; Width = 23; Align = "Center"; NumberFormat = "m/d/yyyy h:mm:ss"; Comment = "EndTime: The End Time of the Meeting. This and all Times are in UTC." }
        @{ Name = "OriginalStartDate"; Width = 15; Align = "Left"; NumberFormat = "m/d/yy"; Comment = "OriginalStartDate: The Original Start Date of the Meeting." }
        @{ Name = "Location"; Width = 10; Align = "Left"; Comment = "Location: The Location of the Meeting." }
        @{ Name = "CalendarItemType"; Width = 15; Align = "Center"; Comment = "CalendarItemType: The Calendar Item Type of the Meeting." }
        @{ Name = "RecurrencePattern"; Width = 20; Align = "Left"; Comment = "RecurrencePattern: The Recurrence Pattern of the Meeting." }
        @{ Name = "AppointmentAuxiliaryFlags"; Width = 30; Align = "Center"; Comment = "AppointmentAuxiliaryFlags: The Appointment Auxiliary Flags of the Meeting." }
        @{ Name = "DisplayAttendeesAll"; Width = 30; Align = "Left"; Comment = "DisplayAttendeesAll: List of the Attendees of the Meeting." }
        @{ Name = "AttendeeCount"; Width = 10; Align = "Center"; Comment = "AttendeeCount: The Attendee Count." }
        @{ Name = "AppointmentState"; Width = 20; Align = "Left"; Comment = "AppointmentState: The Appointment State of the Meeting." }
        @{ Name = "ResponseType"; Width = 10; Align = "Center"; Comment = "ResponseType: The Response Type of the Meeting." }
        @{ Name = "ClientIntent"; Width = 20; Align = "Center"; Comment = "ClientIntent: The Client Intent of the Meeting." }
        @{ Name = "AppointmentRecurring"; Width = 10; Align = "Center"; Comment = "AppointmentRecurring: Is this a Recurring Meeting?" }
        @{ Name = "HasAttachment"; Width = 10; Align = "Center"; Comment = "HasAttachment: Does this Meeting have an Attachment?" }
        @{ Name = "IsCancelled"; Width = 10; Align = "Center"; Comment = "IsCancelled: Is this Meeting Cancelled?" }
        @{ Name = "IsAllDayEvent"; Width = 10; Align = "Center"; Comment = "IsAllDayEvent: Is this an All Day Event?" }
        @{ Name = "Sensitivity"; Width = 12; Align = "Center"; Comment = "Sensitivity: The Sensitivity of the Meeting (Normal, Personal, Private, Confidential)." }
        @{ Name = "IsSeriesCancelled"; Width = 10; Align = "Center"; Comment = "IsSeriesCancelled: Is this a Series Cancelled Meeting?" }
        @{ Name = "SendMeetingMessagesDiagnostics"; Width = 30; Align = "Left"; Comment = "SendMeetingMessagesDiagnostics: Compound Property to describe why meeting was or was not sent to everyone." }
        @{ Name = "AttendeeCollection"; Width = 50; Align = "Left"; Comment = "AttendeeCollection: The Attendee Collection of the Meeting, use -TrackingLogs to get values." }
        @{ Name = "CalendarLogRequestId"; Width = 40; Align = "Center"; Comment = "CalendarLogRequestId: The Calendar Log Request ID of the Meeting." }
        @{ Name = "CleanGlobalObjectId"; Width = 40; Align = "Left"; Comment = "CleanGlobalObjectId: The MeetingID / Clean Global Object ID of the Meeting." }
    )

    foreach ($meta in $headerMeta) {
        $colLetter = $ColumnMap[$meta.Name]
        if ($null -eq $colLetter) { continue }
        $colIdx = [int][char]$colLetter[0] - [int][char]'A' + 1
        # Handle double-letter columns (e.g., "AB", "AM", etc.)
        if ($colLetter.Length -gt 1) {
            $colIdx = 0
            foreach ($c in $colLetter.ToCharArray()) {
                $colIdx = $colIdx * 26 + ([int][char]$c - [int][char]'A' + 1)
            }
        }
        $col = $sheet.Column($colIdx)
        $params = @{ Width = $meta.Width; HorizontalAlignment = $meta.Align }
        if ($meta.ContainsKey("NumberFormat")) { $params.NumberFormat = $meta.NumberFormat }
        $col | Set-ExcelRange @params
        Set-CellComment -Text $meta.Comment -Row $HeaderRow -ColumnNumber $colIdx -Worksheet $sheet
    }

    # Title Row
    $sheet.Row(1) | Set-ExcelRange -HorizontalAlignment Left
    Set-CellComment -Text "For more information see: https://aka.ms/AnalyzeCalLogs"  -Row 1 -ColumnNumber 1  -Worksheet $sheet

    # Set the Header row to be bold and left aligned
    $sheet.Row($HeaderRow) | Set-ExcelRange -Bold -HorizontalAlignment Left
}
}

if ($AllExceptions.IsPresent -and -not $FastExceptions.IsPresent) {
    Write-Warning "-AllExceptions only applies when -FastExceptions is used. The switch will be ignored."
}

# Default to Collecting Tracking Logs (MeetingID only)
if (-not ([string]::IsNullOrEmpty($MeetingID))) {
    if (!$NoTrackingLogs.IsPresent) {
        $TrackingLogs = $true
        Write-Host -ForegroundColor Yellow "Collecting Tracking Logs."
        Write-Host -ForegroundColor Yellow "`tTo skip collecting Tracking Logs, use the -NoTrackingLogs switch."
    } else {
        Write-Host -ForegroundColor Green "Not Collecting Tracking Logs."
    }

    # Default to Collecting Exceptions
    if ((!$NoExceptions.IsPresent) -and ([string]::IsNullOrEmpty($ExceptionDate))) {
        $Exceptions = $true
        Write-Host -ForegroundColor Yellow "Collecting Exceptions."
        Write-Host -ForegroundColor Yellow "`tTo skip collecting Exceptions, use the -NoExceptions switch."
        if ($FastExceptions.IsPresent) {
            Write-Host -ForegroundColor Yellow "`tFast exception collection is enabled and will parse AppointmentRecurrenceBlob first."
            if ($AllExceptions.IsPresent) {
                Write-Host -ForegroundColor Yellow "`tAll Exception dates will be collected."
            } else {
                Write-Host -ForegroundColor Yellow "`tOnly Exception dates from the last 6 months will be collected by default."
            }
        }
    } else {
        Write-Host -ForegroundColor Green "---------------------------------------"
        if ($NoExceptions.IsPresent) {
            Write-Host -ForegroundColor Green "Not Checking for Exceptions"
        } else {
            Write-Host -ForegroundColor Green "Checking for Exceptions on $ExceptionDate"
        }
        Write-Host -ForegroundColor Green "---------------------------------------"
    }
} else {
    # Subject-based search
    $script:SubjectSearch = $true
    if ((!$NoExceptions.IsPresent) -and ([string]::IsNullOrEmpty($ExceptionDate))) {
        $Exceptions = $true
        Write-Host -ForegroundColor Yellow "Using Subject search. Exception collection will run if the search resolves to exactly one MeetingID."
        Write-Host -ForegroundColor Yellow "`tTo skip collecting Exceptions for a single resolved MeetingID, use the -NoExceptions switch."
        if ($FastExceptions.IsPresent) {
            Write-Host -ForegroundColor Yellow "`tFast exception collection is enabled for single-MeetingID Subject results."
            if ($AllExceptions.IsPresent) {
                Write-Host -ForegroundColor Yellow "`tAll Exception dates will be collected."
            } else {
                Write-Host -ForegroundColor Yellow "`tOnly Exception dates from the last 6 months will be collected by default."
            }
        }
    } else {
        Write-Host -ForegroundColor Green "Using Subject search without automatic Exception collection."
        if ($NoExceptions.IsPresent) {
            Write-Host -ForegroundColor Green "`t-NoExceptions was specified, so Exception collection will be skipped."
        } elseif (-not ([string]::IsNullOrEmpty($ExceptionDate))) {
            Write-Host -ForegroundColor Green "`t-ExceptionDate applies only after a specific MeetingID is selected."
        }
    }
    Write-Host -ForegroundColor Yellow "`tTracking Logs still require running the script with -MeetingID."
}

# ===================================================================================================
# Main
# ===================================================================================================

[array]$ValidatedIdentities = CheckIdentities -Identity $Identity

if ($ExportToExcel.IsPresent) {
    CheckExcelModuleInstalled
}

if (-not ([string]::IsNullOrEmpty($Subject)) ) {
    if ($ValidatedIdentities.count -gt 1) {
        Write-Error "Subject-based searches only support a single mailbox, but $($ValidatedIdentities.count) were provided: [$($ValidatedIdentities -join ', ')]."
        Write-Host -ForegroundColor Yellow "Options:"
        Write-Host -ForegroundColor Yellow "  1. Re-run with a single -Identity and -Subject."
        Write-Host -ForegroundColor Yellow "  2. Use -MeetingID instead of -Subject to search multiple mailboxes at once."
        Write-Host -ForegroundColor Yellow "     Tip: Run the script once with -Subject for one user, then use the MeetingID from the output to collect from all participants."
        exit
    }
    $script:Identity = $ValidatedIdentities[0]
    $script:CurrentIdentityRunStartTime = Get-Date
    GetCalLogsWithSubject -Identity $ValidatedIdentities -Subject $Subject
} elseif (-not ([string]::IsNullOrEmpty($MeetingID))) {
    #Validate MeetingID is good
    if ($MeetingID -like "UID:*") {
        $MeetingID = $MeetingID.Replace("UID:", "")
    }
    if ($MeetingID -notlike "040000008*") {
        Write-Error "This does not look like a valid MeetingID: $MeetingID."
        Write-Error "Calendar MeetingID almost always start with 040000008..."
    }
    # Process Logs based off Passed in MeetingID
    foreach ($ID in $ValidatedIdentities) {
        $script:CurrentIdentityRunStartTime = Get-Date
        Write-DashLineBoxColor "Looking for CalLogs from [$ID] with passed in MeetingID."
        Write-Verbose "Running: Get-CalendarDiagnosticObjects -Identity [$ID] -MeetingID [$MeetingID] -CustomPropertyNames $CustomPropertyNameList -WarningAction Ignore -MaxResults $LogLimit -ResultSize $LogLimit -ShouldBindToItem $true;"
        [array] $script:GCDO = GetCalendarDiagnosticObjects -Identity $ID -MeetingID $MeetingID
        $script:Identity = $ID
        if ($script:GCDO.count -gt 0) {
            Write-Host -ForegroundColor Cyan "Found $($script:GCDO.count) CalLogs with MeetingID [$MeetingID]."
            $script:IsOrganizer = (SetIsOrganizer -CalLogs $script:GCDO)
            Write-Host -ForegroundColor Cyan "The user [$ID] $(if ($IsOrganizer) {"IS"} else {"is NOT"}) the Organizer of the meeting."

            $script:IsRoomMB = (SetIsRoom -CalLogs $script:GCDO)
            if ($script:IsRoomMB) {
                Write-Host -ForegroundColor Cyan "The user [$ID] is a Room Mailbox."
            }

            if ($script:IsOrganizer -and (CheckForBifurcation($script:GCDO) -ne $false)) {
                Write-Host -ForegroundColor Red "Warning: No IPM.Appointment found for the Organizer. CalLogs start to expire after 31 days."
            }

            if ($Exceptions.IsPresent) {
                CollectExceptionLogs -Identity $ID -MeetingID $MeetingID
            }

            BuildCSV
            BuildTimeline
        } else {
            Write-Warning "No CalLogs were found for [$ID] with MeetingID [$MeetingID]."
        }
    }
} else {
    Write-Warning "A valid MeetingID was not found, nor Subject. Please confirm the MeetingID or Subject and try again."
}

Write-DashLineBoxColor "Hope this script was helpful in getting and understanding the Calendar Logs.",
"More Info on Getting the logs: https://aka.ms/GetCalLogs",
"and on Analyzing the logs: https://aka.ms/AnalyzeCalLogs",
"If you have issues or suggestion for this script, please send them to: ",
"`t CalLogFormatterDevs@microsoft.com" -Color Yellow -DashChar "="

if ($ExportToExcel.IsPresent) {
    Write-Host
    Write-Host -ForegroundColor Blue -NoNewline "All Calendar Logs are saved to: "
    Write-Host -ForegroundColor Yellow ".\$Filename"
}

# SIG # Begin signature block
# MIIncQYJKoZIhvcNAQcCoIInYjCCJ14CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBNJgrmrTHtig4f
# FsAtLCdoCUrdUnF3VP4PGTbptot5qqCCDMkwggYEMIID7KADAgECAhMzAAACHPrN
# xZvoL37EAAAAAAIcMA0GCSqGSIb3DQEBCwUAMFcxCzAJBgNVBAYTAlVTMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBD
# b2RlIFNpZ25pbmcgUENBIDIwMjQwHhcNMjYwNDE2MTg1OTQxWhcNMjcwNDE1MTg1
# OTQxWjB0MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYD
# VQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQDVsZfgOKmM31HPfoWOoNEiw0SlCiIxUMC0I9NMWbucKOw/e9lP
# oAoehQVu6SG65V4EPzrYsnBnFPNoi4/HoOdjhz1qkrEt4I6tEcxXU6oOeY9zGveC
# /3iBeuhLYxM3M/PkcUoebF+Nednm8OkdSPoDu8imViHPQq/8CQUu0WRR4rE+dMRf
# rpVqfmNi2qWCX94T4MsepijGVkwE//tJg0ryAiYdHT34LSnlG/RSBZmQRGWZ5g8j
# qnKjRParSqMft1gvjuUTVgtWNZfgcLFSK5Wa0myrq8OPcgTGGsRgun+tnSS+IxDT
# xVsAPH1OzvPjwomguByhUe/OcvUN0D5Wmp7xAgMBAAGjggGqMIIBpjAOBgNVHQ8B
# Af8EBAMCB4AwHwYDVR0lBBgwFgYKKwYBBAGCN0wIAQYIKwYBBQUHAwMwHQYDVR0O
# BBYEFNoH7a2YDjOSwpkp6DHcmUS7J+0yMFQGA1UdEQRNMEukSTBHMS0wKwYDVQQL
# EyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxFjAUBgNVBAUT
# DTIzMDAxMis1MDc1NjkwHwYDVR0jBBgwFoAUf1k/VCHarU/vBeXmo9ctBpQSCDEw
# YAYDVR0fBFkwVzBVoFOgUYZPaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9w
# cy9jcmwvTWljcm9zb2Z0JTIwQ29kZSUyMFNpZ25pbmclMjBQQ0ElMjAyMDI0LmNy
# bDBtBggrBgEFBQcBAQRhMF8wXQYIKwYBBQUHMAKGUWh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwQ29kZSUyMFNpZ25pbmcl
# MjBQQ0ElMjAyMDI0LmNydDAMBgNVHRMBAf8EAjAAMA0GCSqGSIb3DQEBCwUAA4IC
# AQAUnEqhaRXe0T3hIJjvdQErEkrA/7bByjn6t5IArODkkRjzkYwtKMc2yYj2quaN
# rLutWw2YZcngKPy1b71YyDJQTy4NDRwaSh9Tw5thrk3NmcPrAHia5vtcBJ1CgtKK
# 7mQbIcQ22d/N3813ayCDDFewu1+jsZmX+r/aTEqaOM4TVxVtRSkuCy8nAXKuChOK
# Li/zA4XuH8iEYqIsj2YoNaeSxVmeGiERXpKdo3dDmYi0kO5w2D8VS4c3+9h6gElY
# BaAAg/dYErBg27qT3vv0zRDJhJufvCNylA8S7/+8H5E/PV5cng6na9VV/w9OV3qu
# uND6zdGa2EX38Glp50F9AIQk3p2xXmcvorDeM4XJ7UlWYBi6g80J1SSOQnInCYFE
# msfUNn3+1AaTJKSJL83quKArTac2pKhu0Yzzzrzo6HrsRiQKzpnRBb1/dMa6P3hz
# 75XbMRBctNsFhZC07WCmjExdLg2eHW5uV0TY8D5+6wozJf7vF3+WHkYPO85Z+BC6
# U4FkNbYNycZ9cE4j1tXRdyDCfml6c0HWPHjNVDObrv9lKt3qUqFpX38VCqVCyNOO
# 1UcXfQiVjJw32U2WUKZjt/neJKHEBsm9kFsLuWzkQ53+qcaSaytmsCnk2gOglrlD
# 5d3kKyvvAw+rzm0lT8K38P6PLxfZQHhu4W8dV7Av8N2ZmDCCBr0wggSloAMCAQIC
# EzMAAAA5O7Y3Gb8GHWcAAAAAADkwDQYJKoZIhvcNAQEMBQAwgYgxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBS
# b290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDExMB4XDTI0MDgwODIwNTQxOFoX
# DTM2MDMyMjIyMTMwNFowVzELMAkGA1UEBhMCVVMxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQ
# Q0EgMjAyNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBANgBnB7jOMeq
# lRYHNa265v4IY9fH8TKhemHfPINe1gpLaV3dhg324WwH06LcHbpnsBukCDNitryo
# 0dtS/EW6I/yEL/bLSY8hKpbfQuWusBPr9qazYcDxCW/qnjb5JsI1s8bNOg3bVATv
# QVL4tcf03aTycsz8QeCdM0l/yHRObJ9QqazM1r6VPEOJ7LL+uEEb73w6QCuhs89a
# 1uv1zerOYMnsneRRwCbpyW11IcggU0cRKDDq1pjVJzIbIF6+oiXXbReOsgeI8zu1
# FyQfK0fVkaya8SmVHQ/tOf23mZ4W9k0Ri22QW9p3UgSC5OUDktKxxcCmGL6tXLfO
# GSWHIIV4YrTJTT6PNty5REojHJuZHArkF9VnHTERWoTjAzfI3kP+5b4alUdhgAZ7
# ttOu1bVnXfHaqPYl2rPs20ji03LOVWsh/radgE17es5hL+t6lV0eVHrVhsssROWJ
# uz2MXMCt7iw7lFPG9LXKGjsmonn2gotGdHIuEg5JnJMJVmixd5LRlkmgYRZKzhxS
# CwyoGIq0PhaA7Y+VPct5pCHkijcIIDm0nlkK+0KyepolcqGm0T/GYQRMhHJlGOOm
# VQop36wUVUYklUy++vDWeEgEo4s7hxN6mIbf2MSIQ/iIfMZgJxC69oukMUXCrOC3
# SkE/xIkgpfl22MM1itkZ35nNXkMolU1lAgMBAAGjggFOMIIBSjAOBgNVHQ8BAf8E
# BAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFH9ZP1Qh2q1P7wXl5qPX
# LQaUEggxMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMA8GA1UdEwEB/wQFMAMB
# Af8wHwYDVR0jBBgwFoAUci06AjGQQ7kUBU7h6qfHMdEjiTQwWgYDVR0fBFMwUTBP
# oE2gS4ZJaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMv
# TWljUm9vQ2VyQXV0MjAxMV8yMDExXzAzXzIyLmNybDBeBggrBgEFBQcBAQRSMFAw
# TgYIKwYBBQUHMAKGQmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMv
# TWljUm9vQ2VyQXV0MjAxMV8yMDExXzAzXzIyLmNydDANBgkqhkiG9w0BAQwFAAOC
# AgEAFJQfOChP7onn6fLIMKrSlN1WYKwDFgAddymOUO3FrM8d7B/W/iQ6DxXsDn7D
# 5W4wMwYeLystcEqfkjz4NURRgazyMu5yRzQh4LqjA4tStTcJh1opExo7nn5PuPBY
# nbu0+THSuVHTe0VTTPVhily/piFrDo3axQ9P4C+Ol5yet+2gTfekICS5xS+cYfSI
# vgn0JksVBVMYVI5QFu/qhnLhsEFEUzG8fvv0hjgkO+lkpV9ty6GkN4vdnd7ya6Q6
# aR9y34aiM1qmxaxBi6OUnyNl6fkuun/diTFnYDLTppOkr/mg5WSfCiDVMNCxtj4w
# PKC5OmHm1DQIt/MNokbbH3UGsFP1QbzsLocuSqLCvH09Io3fDPTmscR9Y75G4qX7
# RTX8AdBPo0I6OEojf39zuFZt0qOHm65YWQE69cZM2ueE1MB05dNNgHK9gTE7zKvK
# /fg8B2qjW88MT/WF5V5uvZGtqa9FSL2RazArA+rDPuf6JGYz4HpgMZHB4S6szWSK
# YBv0VisCzfxgeU+dquXW9bd0auYlOB58DPcOYKdc3Se94g+xL4pcEhbB54JOgAkw
# YTu/9dLeH2pDqeJZAABVDWRQCaXfO5LgyKwKCLYXpigrZYCjUSBcr+Ve8PFWMhVT
# Ql0v4q8J/AUmQN5W4n101cY2L4A7GTQG1h32HHAvfQESWP0xghn+MIIZ+gIBATBu
# MFcxCzAJBgNVBAYTAlVTMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# KDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMjQCEzMAAAIc
# +s3Fm+gvfsQAAAAAAhwwDQYJYIZIAWUDBAIBBQCgga4wGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIAtrVnSoSB6ShBflxArS3qnye27x7GlCDSkCb6SNysYXMEIGCisG
# AQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEBBQAEggEAxjqItBYDT3WQPNJJwhYA
# VVsW0Z9TLfSzWqrzkRPp3aVCm9mN5ICPtBfWXrWvMYsn8naglkZIh77Xxm5fDSA/
# /yOIuy88Ek1OsVHqiMBOH4PQNQ+QvUNeM2FhuHqlBUIGw9ax+Kvwc4aVhzCjJku6
# s+NqtOcyfVVwYHD/NhaH4Z2w9Md1j9nL/jZV/9OFpN+HK856me33TDYjz2Thz1c1
# Y/JpJIaXQHFAMiaje8/4asRnvyBFNtY2KZACSktzPSOjXtarQvUY0dlMthLg+jv+
# SC71EuUoc4tYxJhU4DX35Sl0kCVPAYNKHsadVcJSkkP/UEdNlNlBD2yiSBjfKd7q
# RqGCF7AwghesBgorBgEEAYI3AwMBMYIXnDCCF5gGCSqGSIb3DQEHAqCCF4kwgheF
# AgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFaBgsqhkiG9w0BCRABBKCCAUkEggFFMIIB
# QQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCCHwDNceAkRbQeXnwFb
# 1MXQCkrXYQWSaBcTRnKf2Fh5dAIGaeu6CCHsGBMyMDI2MDUxOTE1MDYxMC4xMjla
# MASAAgH0oIHZpIHWMIHTMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0
# ZWQxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo0MDFBLTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaCCEf4wggcoMIIFEKAD
# AgECAhMzAAACGV6y2FR19LGNAAEAAAIZMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29m
# dCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTI1MDgxNDE4NDgyNloXDTI2MTExMzE4
# NDgyNlowgdMxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTAr
# BgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUG
# A1UECxMeblNoaWVsZCBUU1MgRVNOOjQwMUEtMDVFMC1EOTQ3MSUwIwYDVQQDExxN
# aWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkqhkiG9w0BAQEFAAOC
# Ag8AMIICCgKCAgEApqFIyUkzIyxpL3Q03WmLuy4G9YIUScznhKr+cHOT+/u7Parx
# I96gxxb1WrWuAxB8qjGLfsbImx8V3ouK1nUcf+R/nsnXas5/iTgV/Tl3QTRGT0De
# uXBNbpHqc+wC1NiTyA76gLnirvSBEoBzlrpNQFEnuwdbPLCLpTS3KWSCu5J02b+R
# FWR/kcFzVxnhoE3gIaeURtrGKGBZGKLBXvqggkDENtKkvtvRT32xLvAvL/RpReu5
# z18ZojCs72ZSoa74Dy8YbaWsDm3OZOpJRZxZsPKCHZ6xNqgFKf0xNHj0t9v0Q3W+
# 2z5gAVaasJJCvR52Sl0XJ2AOf3l0LSetXgUA5gD5IQ1RvEslTmNnSouTrGID3D1n
# jY7mBu0puiIdPK2jK/1Weef2+YR4cQpWQkeBZmXidh9AuWdlwxKQL15LJ6K2dw8y
# /t/PBhmLyt6QAf0CepWRdgZnMytVAUuWHwlZRV9JLY7aX8D55eL9+cOLpX3bGNOm
# N24UpIW8qtZaqXaesFvIOW23JNLhaaQVvObr1eu7GE/5Mn43e+/DbtdYl/bLP2IQ
# 1xYEJdSbcUkDFfW3KlZEh+nBKDtaRnNRkbgIgxIbKdT38OKQwZ/aA4uSsiAg6nEP
# iWBHGuytIo5wU75M5VdjhEqqTHfXYu8BJi6GTzvWT+9ekfMXezqCkksxaG8CAwEA
# AaOCAUkwggFFMB0GA1UdDgQWBBSAaOo5HWatNzqZn1IF1fcD6nr3ITAfBgNVHSME
# GDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSgUqBQhk5odHRw
# Oi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBUaW1l
# LVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBeMFwGCCsG
# AQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01p
# Y3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAMBgNVHRMB
# Af8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDAN
# BgkqhkiG9w0BAQsFAAOCAgEAXxzVZLLXBFfoCCCTiY7MHXdb7civJSTfrHYJC5Ok
# 2NN75NpzTMT9V2TcIQjfQ3AFUbh1NBAYtMUuwxC6D4ceEXG5lXAnbvkC9YjeLVDR
# yImXYYmft7z+Qpl9t3C/8a0tiqnOz8Ue8/DYLtMTgvWMnsqLNjILDaImOfnHI36T
# LCjGFe8RYLXGdCUdOLlfAdMGePxSTA3TAAOc+GQbmPWjrguLWbxvnl3NVjRvrBZV
# kxFMoVZH0f7qGwDOShjpnv5nYnQ48ufL0uBz52RbPGdX4Fv9+UGOrBprmcHzmIut
# FtJec2Y4kujNtTK2wBGgWscEOVhFiaVdje8VLJ7MVNKE5TmsuGM3jTLr1nuR5AFG
# s3UKkP7g3cQD4cHK7XdLiTm7e606QJ+WqeQsADYE9dvU9wIUbI9Dl4UcIErFw+FH
# aWSTrkfJ4SvLmhKnl5khhpJ1sF3z6e1BxepUliXHqzRLiHWihWIWESF8IHElF3PO
# xbP4VJqHBiYvaXMV0SyRgwoD6zXddbUnX9WR6JL2BlqAjjHxINwelsp/VhxAWThz
# uMA58LxvE/VAzjfFF4Wm7a1ZALmJVw3oL/s/uxo1Op4tcT+hfZ9uN1htC1JN4DuR
# qFfLttjuoAmUQobO5zUFRzvCn8Ck/hiO+bzR15sqkjlxLMyMjpkc/ef4SUUikD46
# 8vUwggdxMIIFWaADAgECAhMzAAAAFcXna54Cm0mZAAAAAAAVMA0GCSqGSIb3DQEB
# CwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMTIwMAYD
# VQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAxMDAe
# Fw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMyMjVaMHwxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0
# YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA5OGm
# TOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51yMo1V/YBf2xK4OK9uT4XYDP/XE/H
# ZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY6GB9alKDRLemjkZrBxTzxXb1hlDc
# wUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9cmmvHaus9ja+NSZk2pg7uhp7M62A
# W36MEBydUv626GIl3GoPz130/o5Tz9bshVZN7928jaTjkY+yOSxRnOlwaQ3KNi1w
# jjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDuaRr3tpK56KTesy+uDRedGbsoy1cCG
# MFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74kpEeHT39IM9zfUGaRnXNxF803RKJ
# 1v2lIH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2K26oElHovwUDo9Fzpk03dJQcNIIP
# 8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5TI4CvEJoLhDqhFFG4tG9ahhaYQFz
# ymeiXtcodgLiMxhy16cg8ML6EgrXY28MyTZki1ugpoMhXV8wdJGUlNi5UPkLiWHz
# NgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9QBXpsxREdcu+N+VLEhReTwDwV2xo3
# xwgVGD94q0W29R6HXtqPnhZyacaue7e3PmriLq0CAwEAAaOCAd0wggHZMBIGCSsG
# AQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUCBBYEFCqnUv5kxJq+gpE8RjUpzxD/
# LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJlpxtTNRnpcjBcBgNVHSAEVTBTMFEG
# DCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wEwYDVR0lBAwwCgYIKwYB
# BQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8G
# A1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQw
# VgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9j
# cmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUF
# BwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwDQYJKoZIhvcNAQEL
# BQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/ypb+pcFLY+TkdkeLEGk5c9MTO1OdfC
# cTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulmZzpTTd2YurYeeNg2LpypglYAA7AF
# vonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM9W0jVOR4U3UkV7ndn/OOPcbzaN9l
# 9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECWOKz3+SmJw7wXsFSFQrP8DJ6LGYnn
# 8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4FOmRsqlb30mjdAy87JGA0j3mSj5m
# O0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3UwxTSwethQ/gpY3UA8x1RtnWN0SCyx
# TkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPXfx5bRAGOWhmRaw2fpCjcZxkoJLo4
# S5pu+yFUa2pFEUep8beuyOiJXk+d0tBMdrVXVAmxaQFEfnyhYWxz/gq77EFmPWn9
# y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGConsXHRWJjXD+57XQKBqJC4822rpM
# +Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdCPSWU5nR0W2rRnj7tfqAxM328y+l7vzhw
# RNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEGahC0HVUzWLOhcGbyoYIDWTCCAkEC
# AQEwggEBoYHZpIHWMIHTMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0
# ZWQxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo0MDFBLTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIa
# AxUAMXYp/Wqqdyb0enigrLfxl0InAz6ggYMwgYCkfjB8MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1T
# dGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQsFAAIFAO22g5MwIhgPMjAyNjA1MTkw
# NjM3MDdaGA8yMDI2MDUyMDA2MzcwN1owdzA9BgorBgEEAYRZCgQBMS8wLTAKAgUA
# 7baDkwIBADAKAgEAAgIJsgIB/zAHAgEAAgITVDAKAgUA7bfVEwIBADA2BgorBgEE
# AYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYag
# MA0GCSqGSIb3DQEBCwUAA4IBAQAKWHnVsB/QLLPjwEzN6DFXsaGBjVlE9I966Mv4
# Yb9WNAhSzhdNPsDi6E3nAch0B6eARKBk+3xLm/Sq109dKX4KaZJnlUXoKcyt6e2n
# oZafHAVtufRq9KMnlq846V261SXxEDZ50wLgkr820sEQmfBwNg14exzyg+sMv6Cn
# ah5yRj71mtHC9t+HBC+ani4BMGzm0EC5ZUoMPHK2ufMEW5fYifHPEMVUD7R4Zd9+
# GSIhYKedG6Kxq3vhHS9gmj4YHHB2uwoFzANLmfiJ1UbGc2r/beEF14E3JfrOfz3e
# rHso9jMnvQEEeey36Dy1V+ek5cp8XrC6vjT52HlX9VWybptAMYIEDTCCBAkCAQEw
# gZMwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UE
# AxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAIZXrLYVHX0sY0A
# AQAAAhkwDQYJYIZIAWUDBAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0B
# CRABBDAvBgkqhkiG9w0BCQQxIgQghCtDvEpi9+ChCT1WRzkGC3fwlamXjElPkCu3
# MgXDE4EwgfoGCyqGSIb3DQEJEAIvMYHqMIHnMIHkMIG9BCDckX633E1y1EF32V18
# zQcrsgjzI9+3Le7mlvk2OebthjCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwAhMzAAACGV6y2FR19LGNAAEAAAIZMCIEIK0IBBFnYBy6xrI/sGEY
# A17SiX9DGoYc3aa3tv+5Jsa7MA0GCSqGSIb3DQEBCwUABIICAC8Hi//vvbLx9VzP
# bb7+I6oJW0UDCnIvVXmehKlfCHpUooKozT4GvS6HxvnCjwIi7MzAC2PVtAmi3NpV
# zzKDiOOajy5c6rwapH9x0Kr7iaDs5TaCflkgg+Y5DuwFBbMnqZuEUjg89BBw1Oob
# U5b/KDznPLv1D+HXWlg43GUli7bEkc2Ddb8/wFpEOT9eTrLy8D2tueWYTxsW8sC5
# /SMA82ZXqzqOE3tgEFhtxxxyPi+655ZBrhFvLYq0ZGL3ZuC2sikU2waoJtkoTlCd
# 6MLZc3jJigh5RYgkRyM0EiQl0u2SM+SukxFlLJh1yA9qxd4SwDCp3WmNrKGMpXJB
# af4/mNa/bgGupBiecj3FDrlOdyC+VLfkqVhW7yGcGzzF2c7oiCqMttmOSkTGiFw+
# 9CQgoRIzpbaTP7WQUttZnsTemKRKcI09aHc0tpH4VgoAFDIZJSB2p9mTfwsoo/f2
# GlnYbub5POFdc6uccW/3BwfAds7pepWnaXfGTJE2GkrDo6sZmy53p/VTjoVHlPh6
# ntPZLj5YO71vHXPL70+P6yv5cyLtBHba51rOIjVQ9s07ynsWL+6y/KgmpHeti26h
# mSTGj36mzIq/O/kDcZ3qF0R6aolthFsYogGhm9AuphkUCmGbdh8EtQGzA44ovZom
# fD1/qyw11G9Vg2ZV5e/9JOe96bvj
# SIG # End signature block
