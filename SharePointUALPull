# Modify the values for the following variables to configure the audit log search.
[DateTime]$start = "2023-07-01"
[DateTime]$end = "2023-10-08"
$record = "SharePointFileOperation"
$resultSize = 5000
$intervalMinutes = 320

# Load siteids from file
$siteids = Get-Content "C:\Users\<user>\Documents\Tools\PScripts\SiteIDs.txt"

# Start script
Function Write-LogFile ([String]$Message) {
    $final = [DateTime]::Now.ToUniversalTime().ToString("s") + ":" + $Message
    $final | Out-File $logFile -Append
}

foreach ($siteid in $siteids) {
    $logFile = "C:\Users\<user>\Documents\Tools\PScripts\AuditLogSearchLog - " + $siteid + ".csv"
    $outputFile = "C:\Users\<user>\Documents\Tools\PScripts\AuditLogRecords - " + $siteid + ".csv"
    [DateTime]$currentStart = $start
    [DateTime]$currentEnd = $end

    Write-LogFile "BEGIN: Retrieving audit records between $($start) and $($end), RecordType=$record, PageSize=$resultSize, SiteID=$siteid."
    Write-Host "Retrieving audit records for the date range between $($start) and $($end), RecordType=$record, ResultsSize=$resultSize, SiteID=$siteid"

    $totalCount = 0
    while ($true) {
        $currentEnd = $currentStart.AddMinutes($intervalMinutes)
        if ($currentEnd -gt $end) {
            $currentEnd = $end
        }

        if ($currentStart -eq $currentEnd) {
            break
        }

        $sessionID = [Guid]::NewGuid().ToString() + "_" +  "ExtractLogs" + (Get-Date).ToString("yyyyMMddHHmmssfff")
        Write-LogFile "INFO: Retrieving audit records for activities performed between $($currentStart) and $($currentEnd)"
        Write-Host "Retrieving audit records for activities performed between $($currentStart) and $($currentEnd)"
        $currentCount = 0

        $sw = [Diagnostics.StopWatch]::StartNew()
        do {
            $results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -RecordType $record -SessionId $sessionID -Siteids $siteid -SessionCommand ReturnLargeSet -ResultSize $resultSize

            if (($results | Measure-Object).Count -ne 0) {
                $results | export-csv -Path $outputFile -Append -NoTypeInformation

                $currentTotal = $results[0].ResultCount
                $totalCount += $results.Count
                $currentCount += $results.Count
                Write-LogFile "INFO: Retrieved $($currentCount) audit records out of the total $($currentTotal)"

                if ($currentTotal -eq $results[$results.Count - 1].ResultIndex) {
                    $message = "INFO: Successfully retrieved $($currentTotal) audit records for the current time range. Moving on!"
                    Write-LogFile $message
                    Write-Host "Successfully retrieved $($currentTotal) audit records for the current time range. Moving on to the next interval." -foregroundColor Yellow
                    ""
                    break
                }
            }
        } while (($results | Measure-Object).Count -ne 0)

        $currentStart = $currentEnd
    }

    Write-LogFile "END: Retrieving audit records between $($start) and $($end), RecordType=$record, PageSize=$resultSize, SiteID=$siteid, total count: $totalCount."
    Write-Host "Script complete! Finished retrieving audit records for the date range between $($start) and $($end). Total count: $totalCount" -foregroundColor Green
}
