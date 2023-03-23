#Clear any potentially uninitialized variable
$readFileOrManualEntry = $outputMode = $fileNotFound = $hostsFilePath = $null

#Initialize lists for input, results (output), and errors.
$HostsForLookup = New-Object System.Collections.Generic.List[System.Object]
$results = New-Object System.Collections.Generic.List[System.Object]
$errors = New-Object System.Collections.Generic.List[System.Object]

write-output "`n"
write-output "`t`t`t`t`t   *!*!* Trace Route *!*!*`n"

#Determine input mode.
do {
    $readFileOrManualEntry = read-host -prompt "Read Input From File (1) or Manual Entry (2) [Default = Read Input From File]"
    if (!$readFileOrManualEntry) { $readFileOrManualEntry = 1 }
} 
while ($readFileOrManualEntry -ne 1 -and $readFileOrManualEntry -ne 2 -and $readFileOrManualEntry -ne "Q")
if ($readFileOrManualEntry -eq "Q") { exit }

#Read ips in if input mode = 1 (i.e. input from file).
if ($readFileOrManualEntry -eq 1) {
    do {
        $hostsFilePath = read-host -prompt "`nHostname/IP Address Input File [Default=.\HostsForLookup.txt]" 
        if(!$hostsFilePath) { $hostsFilePath = ".\HostsForLookup.txt" }
        if ($hostsFilePath -ne "Q") { 
            $fileNotFound = $(!$(test-path $hostsFilePath -PathType Leaf))
            if ($fileNotFound) { write-output "`n`tFile '$hostsFilePath' Not Found or Path Specified is a Directory!`n" }
        }
        if($fileNotFound) {
            write-output "`n** Remember To Enter Fully Qualified Filenames If Files Are Not In Current Directory **" 
            write-output "`n`tFile must contain one ip address per line.`n"
        }
    }
    while ($fileNotFound -and $hostsFilePath -ne "Q")
    if ($hostsFilePath -eq "Q") { exit }

    $HostsForLookup = Get-Content $hostsFilePath -ErrorAction Stop
}
#Prompt for ips if input mode = 2 (i.e. manual entry).
else {
    $hostCount = 0
    write-output "`n`nEnter 'f' once finished. Minimum 1 entry. (Enter 'q' to exit.)`n"
    do {
        $hostInput = read-host -prompt "Hostname/IP Address ($($hostCount + 1))"
        if ($hostInput -ne "F" -and $hostInput -ne "B" -and $hostInput -ne "Q" -and 
            ![string]::IsNullOrEmpty($hostInput)) {
            if ($hostInput -eq 'localhost') { $hostInput = $ENV:Computername }
            $HostsForLookup.Add($hostInput)
            $hostCount++
            }
    }
    while (($hostInput -ne "F" -and $hostInput -ne "B" -and $hostInput -ne "Q") -or 
            ($hostCount -lt 1 -and $hostInput -ne "B" -and $hostInput -ne "Q"))

    if ($hostInput -eq "Q") { exit }
}

#Determine output mode.
do { 
    $outputMode = read-host -prompt "`nSave To File (1), Console Output (2), or Both (3) [Default=3]"
    if (!$outputMode) { $outputMode = 3 }
}
while ($outputMode -ne 1 -and $outputMode -ne 2 -and $outputMode -ne 3 -and $outputMode -ne "Q")
if ($outputMode -eq "Q") { exit }

#If file output selected, determine location/filename...
if ($outputMode -eq 1 -or $outputMode -eq 3) {
        write-output "`n* To save to any directory other than the current, enter fully qualified path name. *"
        write-output   "*              Leave this entry blank to use the default file name of               *"
        write-output   "*                        '$defaultOutFileName',                          *"
        write-output   "*                which will save to the current working directory.                  *"
        write-output   "*                                                                                   *"
        write-output   "*  THE '.csv' EXTENSION WILL BE APPENDED AUTOMATICALLY TO THE FILENAME SPECIFIED.   *`n"

    $defaultOutFileName = "TraceRouteOutput-$(Get-Date -Format MMddyyyy_HHmmss)"

    do { 
        $outputFileName = read-host -prompt "Save As [Default=$defaultOutFileName]"
        if ($outputFileName -eq "Q") { exit }
        if(!$outputFileName) { $outputFileName = $defaultOutFileName }
        $pathIsValid = $true
        $overwriteConfirmed = "Y"
        $outputFileName += ".csv"
        #Test for valid file name and check if file already exists...                                
        $pathIsValid = Test-Path -Path $outputFileName -IsValid
        if ($pathIsValid) {          
            $fileAlreadyExists = Test-Path -Path $outputFileName
            if ($fileAlreadyExists) {
                do {
                    $overWriteConfirmed = read-host -prompt "File '$outputFileName' Already Exists. Overwrite (Y) or Cancel (N)"       
                    if ($overWriteConfirmed -eq "Q") { exit }
                } while ($overWriteConfirmed -ne "Y" -and $overWriteConfirmed -ne "N")
            }
        }

        else { 
            write-output "* Path is not valid. Try again. ('q' to quit.) *"
        }
    }
    while (!$pathIsValid -or $overWriteConfirmed -eq "N")
}

#Process lookup...
$HostsForLookup | ForEach-Object {
    $thisHost = $_
    Try { 
        $thisResult = Test-NetConnection -TraceRoute $thisHost -InformationLevel Detailed -ErrorAction Stop -WarningAction Stop
        $count=1
        $thisResult.TraceRoute | ForEach-Object {
            $results.Add([PSCustomObject]@{'Hostname/IP Address'=$thisHost;
                                            'Hop Number'=$count++;
                                            'Hop Address'= "$_"})
        }
    }
    Catch { $errors.Add([PSCustomObject]@{'Hostname/IP Address'=$thisHost;
                                          'Error Message'= $(if ($_.Exception.Message -like "*WarningPreference*") {
                                                                 $_.Exception.Message.split(':')[1..$_.Exception.Message.split(':').Length] -join ":"
                                                             }
                                                             else { $_.Exception.Message }
                                                             )
                                         })
    }
}

if ($outputMode -eq 1 -or $outputMode -eq 3) {
    if ($results) { $results | Export-CSV -Path $outputFileName -NoTypeInformation }
    if ($errors) {
        Add-Content -Path $outputFileName -Value "`r`n** Errors **"
        $errors | Select-Object | ConvertTo-CSV -NoTypeInformation | Add-Content -Path $outputFileName
    }
}
if ($outputMode -eq 2 -or $outputMode -eq 3) {
    if ($results) { 
        write-Output "`n`t`t`t*** Results ***"
        $results | Format-Table
    }
    if ($errors) {
        write-Output "`t`t`t*** Errors ***"
        $errors | Format-Table
    }
}

if($outputMode -eq 1) {
    write-host "`nTask Complete. Press enter to exit..." -NoNewLine
    $Host.UI.ReadLine()
}
else { 
    write-host "Task Complete. Press enter to exit..." -NoNewLine
    $Host.UI.ReadLine()
}

#References:
# https://stackoverflow.com/questions/44397795/dns-name-from-ip-address