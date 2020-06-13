<#
Name - High Severity Incident Report
Author - Ayush Lal
Description - StatusHub Report that will be automatically emailed to a specific Distribution List to provide information on any open High Severity Incidents.
###############################################################################>

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$sh_service_url = "[STATUSHUB URL WITH API KEY]"
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/json")
$headers.Add("Accept", "application/vnd.statushub.v2")

$services = Invoke-RestMethod -Uri $sh_service_url -Method GET -Headers $headers

# Creating global variables
$SH_report = $services | Select-Object title, start_time, incident_updates, incident_updates.body, id

$highsev_count = $SH_report.count

$getSevA_count = $SH_report.where{ $_.title -match 'SevA' }
$SevA_count = $getSevA_count.count

$getSevB_count = $SH_report.where{ $_.title -match 'SevB' }
$SevB_count = $getSevB_count.count

$getSevC_count = $SH_report.where{ $_.title -match 'SevC' }
$SevC_count = $getSevC_count.count

$getRegA_count = $SH_report.where{ $_.title -match 'RegA' }
$RegA_count = $getRegA_count.count

$getRegB_count = $SH_report.where{ $_.title -match 'RegB' }
$RegB_count = $getRegB_count.count

$getBCC_count = $SH_report.where{ $_.title -match 'BCC' }
$BCC_count = $getBCC_count.count

$getOpA_count = $SH_report.where{ $_.title -match 'OpA' }
$OpA_count = $getOpA_count.count


<# Begin Open SH Report
----------------------------------------------------#>
foreach ($entry in $SH_report) {
    $ti = $entry.title
    $inc = $entry.incident_updates.body | Select-String -Pattern 'INC(\d+)' -AllMatches | ForEach-Object { $_.matches } | ForEach-Object { $_.value } | Select-Object -Unique
    $incident_update = $entry.incident_updates.body[0]
    $incident_update_inital = $entry.incident_updates.body
    [datetime]$dt = $entry.start_time
    $DateTime = $dt.ToString("dd/MM/yyyy")
    [datetime]$endDate = $entry.start_time
    $endDate2 = $endDate.AddDays(10)
    $endDate3 = $endDate2.ToString("dd/MM/yyyy")
    $today_date = Get-Date
    $today_date2 = $today_date.ToString("dd/MM/yyyy")
    $INC_link = "[REMEDY TICKET LINK]"
    $SH_commsLink = "https://statushub.io/incidents/" + $entry.id
    $getIC = $entry.incident_updates.body | select-string -Pattern '(Incident Commander:\s+\w+\s+\w+)' -AllMatches | ForEach-Object { $_.matches } | ForEach-Object { $_.value } | Select-Object -Unique
    $IC = $getIC.split("Incident Commander: ")

    if ($today_date2 -ge $endDate3) {
        Write-Host "mark as red"
    }
    else {
        write-host "dont mark as red"
    }
    
    if ($ti -match "SevA") {
        if ($incident_update_inital.count -lt 2) {
            if ($today_date2 -gt $endDate3) {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td class='textred'>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update_inital)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
            else {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update_inital)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
        }
        else {
            if ($today_date2 -gt $endDate3) {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td class='textred'>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
            else {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
        }
    }

    elseif ($ti -match "SevB") {
        if ($incident_update_inital.count -lt 2) {
            if ($today_date2 -gt $endDate3) {
                if ($inc -like "*INC*") {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
                else {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter textred'>No Ticket/Alert</td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
            }
            else {
                if ($inc -like "*INC*") {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
                else {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter textred'>No Ticket/Alert</td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
            }
        }
        else {
            if ($today_date2 -gt $endDate3) {
                if ($inc -like "*INC*") {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
                else {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter textred'>No Ticket/Alert</td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
            }
            else {
                if ($inc -like "*INC*") {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
                else {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter textred'>No Ticket/Alert</td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
            }
        }
    }

    elseif ($ti -match "SevC") {
        if ($incident_update_inital.count -lt 2) {
            if ($today_date2 -gt $endDate3) {
                if ($inc -like "*INC*") {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
                else {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter textred'>No Ticket/Alert</td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
            }
            else {
                if ($inc -like "*INC*") {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
                else {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter textred'>No Ticket/Alert</td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
            }
        }
        else {
            if ($today_date2 -gt $endDate3) {
                if ($inc -like "*INC*") {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
                else {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter textred'>No Ticket/Alert</td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
            }
            else {
                if ($inc -like "*INC*") {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
                else {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($IC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter textred'>No Ticket/Alert</td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
            }
        }
    }

    elseif ($ti -match "RegA") {
        if ($incident_update_inital.count -lt 2) {
            if ($today_date2 -gt $endDate3) {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td class='textred'>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update_inital)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
            else {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update_inital)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
        }
        else {
            if ($today_date2 -gt $endDate3) {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td class='textred'>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
            else {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
        }
    }

    elseif ($ti -match "RegB") {
        if ($incident_update_inital.count -lt 2) {
            if ($today_date2 -gt $endDate3) {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td class='textred'>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update_inital)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
            else {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update_inital)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
        }
        else {
            if ($today_date2 -gt $endDate3) {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td class='textred'>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
            else {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
        }
    }

    elseif ($ti -match "OpA") {
        if ($incident_update_inital.count -lt 2) {
            if ($today_date2 -gt $endDate3) {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td class='textred'>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update_inital)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
            else {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update_inital)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
        }
        else {
            if ($today_date2 -gt $endDate3) {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td class='textred'>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
            else {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
        }
    }

    elseif ($ti -match "HighSev" -or "High Sev") {
        if ($incident_update_inital.count -lt 2) {
            if ($today_date2 -gt $endDate3) {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td class='textred'>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update_inital)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
            else {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update_inital)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
        }
        else {
            if ($today_date2 -gt $endDate3) {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td class='textred'>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
            else {
                $arr += "
                <tr>
                <td class='textleft'>$($ti)</td>
                <td class='textcenter'>$($IC)</td>
                <td>$($DateTime)</td>
                <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                <!--<td class='textleft'>$($incident_update)</td>-->
                <td class='textcenter'><a href='$SH_commsLink'>View</td>
                </tr>`r`n
                "
            }
        }
    }

    elseif ($ti -match "BCC") {
        if ($incident_update_inital.count -lt 2) {
            if ($today_date2 -gt $endDate3) {
                if ($inc -like "*INC*") {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($getIC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
                else {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($getIC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='tdblack'></td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
            }
            else {
                if ($inc -like "*INC*") {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($getIC)</td>
                    <td>$($DateTime)</td>
                    <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
                else {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($getIC)</td>
                    <td>$($DateTime)</td>
                    <td class='tdblack'></td>
                    <!--<td class='textleft'>$($incident_update_inital)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
            }
        }
        else {
            if ($today_date2 -gt $endDate3) {
                if ($inc -like "*INC*") {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($getIC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                    <!--<td class='textleft'>$($incident_update)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
                else {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($getIC)</td>
                    <td class='textred'>$($DateTime)</td>
                    <td class='tdblack'></td>
                    <!--<td class='textleft'>$($incident_update)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
            }
            else {
                if ($inc -like "*INC*") {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($getIC)</td>
                    <td>$($DateTime)</td>
                    <td class='textcenter'><a href='$INC_link'>$($inc)</a></td>
                    <!--<td class='textleft'>$($incident_update)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
                else {
                    $arr += "
                    <tr>
                    <td class='textleft'>$($ti)</td>
                    <td class='textcenter'>$($getIC)</td>
                    <td>$($DateTime)</td>
                    <td class='tdblack'></td>
                    <!--<td class='textleft'>$($incident_update)</td>-->
                    <td class='textcenter'><a href='$SH_commsLink'>View</td>
                    </tr>`r`n
                    "
                }
            }
        }
    }

    else {
        $arr += "
        <tr>
        <td class='textleft'>$($ti)</td>
        <td class='textcenter'>$($IC)</td>
        <td>$($inc)</td>
        <!--<td>$($incident_update)</td>-->
        <td class='textcenter'><a href='$SH_commsLink'>View</td>
        </tr>`r`n
        "
    }


    $openArray = $arr
}


<# Begin Resolved SH Report
----------------------------------------------------#>
$todayDate = (Get-Date).AddDays(-1)
$previousDay = $todayDate.ToString("yyyy-MM-dd")

$resolvedSH_URL = "[STATUSHUB_URL_WITH_APIKEY]$previousDay"
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/json")
$headers.Add("Accept", "application/vnd.statushub.v2")

$Rservices = Invoke-RestMethod -Uri $resolvedSH_URL -Method GET -Headers $headers
$resolvedSH = $Rservices | Select-Object title, start_time, incident_updates, incident_updates.body, id

# Gets an overall count of open incidents (used later in the HTML body)
$res_highsev_count = $resolvedSH.count

foreach ($item in $resolvedSH) {
    $resolved_incident_title = $item.Title
    $resolved_INC = $item.incident_updates.body | Select-String -Pattern 'INC(\d+)' -AllMatches | ForEach-Object { $_.matches } | ForEach-Object { $_.value } | Select-Object -Unique
    $resolved_REQ = $item.incident_updates.body | Select-String -Pattern 'REQ(\d+)' -AllMatches | ForEach-Object { $_.matches } | ForEach-Object { $_.value } | Select-Object -Unique
    $resolved_PD = $item.incident_updates.body | Select-String -Pattern '#(\d+)' -AllMatches | ForEach-Object { $_.matches } | ForEach-Object { $_.value } | Select-Object -Unique
    $resolved_PD2 = $resolved_PD -replace '[#]'
    $incident_update = $item.incident_updates.body[0]
    [datetime]$Date = $item.start_time
    $INC_date = ($Date.ToString("dd/MM/yyyy"))
    $rINC_link = "[REMEDY LINK]"
    $rPD_link = "https://pagerduty.com/incidents/$resolved_PD2"
    $res_SH_commsLink = "https://statushub.io/incidents/" + $item.id


    if ($resolved_incident_title -match "SevA") {
        if ($resolved_INC -like "*INC*") {
            $rArr += "
            <tr class='SevA'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rINC_link'>$($resolved_INC)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_REQ -like "*REQ*") {
            $rArr += "
            <tr class='SevA'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td>$($resolved_REQ)</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_PD -like "*#*") {
            $rArr += "
            <tr class='SevA'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rPD_link'>PD $($resolved_PD)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        else {
            $rArr += "
            <tr class='SevA'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td class='textred textcenter'>No Ticket/Alert</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
    }

    elseif ($resolved_incident_title -match "SevB") {
        if ($resolved_INC -like "*INC*") {
            $rArr += "
            <tr class='SevB'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rINC_link'>$($resolved_INC)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_REQ -like "*REQ*") {
            $rArr += "
            <tr class='SevB'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td>$($resolved_REQ)</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_PD -like "*#*") {
            $rArr += "
            <tr class='SevB'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rPD_link'>PD $($resolved_PD)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        else {
            $rArr += "
            <tr class='SevB'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td class='textred textcenter'>No Ticket/Alert</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
    }

    elseif ($resolved_incident_title -match "SevC") {
        if ($resolved_INC -like "*INC*") {
            $rArr += "
            <tr class='SevB'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rINC_link'>$($resolved_INC)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_REQ -like "*REQ*") {
            $rArr += "
            <tr class='SevB'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td>$($resolved_REQ)</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_PD -like "*#*") {
            $rArr += "
            <tr class='SevB'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rPD_link'>PD $($resolved_PD)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        else {
            $rArr += "
            <tr class='SevB'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td class='textred textcenter'>No Ticket/Alert</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
    }

    elseif ($resolved_incident_title -match "RegA") {
        if ($resolved_INC -like "*INC*") {
            $rArr += "
            <tr class='RegA'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rINC_link'>$($resolved_INC)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_REQ -like "*REQ*") {
            $rArr += "
            <tr class='RegA'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td>$($resolved_REQ)</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_PD -like "*#*") {
            $rArr += "
            <tr class='RegA'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rPD_link'>PD $($resolved_PD)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        else {
            $rArr += "
            <tr class='RegA'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td class='textred textcenter'>No Ticket/Alert</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
    }

    elseif ($resolved_incident_title -match "RegB") {
        if ($resolved_INC -like "*INC*") {
            $rArr += "
            <tr class='RegB'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rINC_link'>$($resolved_INC)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_REQ -like "*REQ*") {
            $rArr += "
            <tr class='RegB'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td>$($resolved_REQ)</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_PD -like "*#*") {
            $rArr += "
            <tr class='RegB'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rPD_link'>PD $($resolved_PD)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        else {
            $rArr += "
            <tr class='RegB'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td class='textred textcenter'>No Ticket/Alert</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
    }

    elseif ($resolved_incident_title -match "OpA") {
        if ($resolved_INC -like "*INC*") {
            $rArr += "
            <tr class='OpA'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rINC_link'>$($resolved_INC)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_REQ -like "*REQ*") {
            $rArr += "
            <tr class='OpA'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td>$($resolved_REQ)</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_PD -like "*#*") {
            $rArr += "
            <tr class='OpA'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rPD_link'>PD $($resolved_PD)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        else {
            $rArr += "
            <tr class='OpA'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td class='textred textcenter'>No Ticket/Alert</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
    }

    elseif ($resolved_incident_title -match "BCC") {
        if ($resolved_INC -like "*INC*") {
            $rArr += "
            <tr class='BCC'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td class='textcenter'><a href ='$rINC_link'>$($resolved_INC)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_REQ -like "*REQ*") {
            $rArr += "
            <tr class='BCC'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td>$($resolved_REQ)</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_PD -like "*#*") {
            $rArr += "
            <tr class='BCC'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rPD_link'>PD $($resolved_PD)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        else {
            $rArr += "
            <tr class='BCC'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td class='textred textcenter'>No Ticket/Alert</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
    }

    elseif ($resolved_incident_title -match "HighSev" -or "High Sev") {
        if ($resolved_INC -like "*INC*") {
            $rArr += "
            <tr class='HighSev'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rINC_link'>$($resolved_INC)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_REQ -like "*REQ*") {
            $rArr += "
            <tr class='HighSev'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td>$($resolved_REQ)</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        elseif ($resolved_PD -like "*#*") {
            $rArr += "
            <tr class='HighSev'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td><a href ='$rPD_link'>PD $($resolved_PD)</a></td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
        else {
            $rArr += "
            <tr class='HighSev'>
            <td class='textleft'>$($resolved_incident_title)</td>
            <td>$($INC_date)</td>
            <td class='textred textcenter'>No Ticket/Alert</td>
            <!--<td class='textleft'>$($incident_update)</td>-->
            <td class='textcenter'><a href='$res_SH_commsLink'>View</td>
            </tr>`r`n
            "
        }
    }

    else {
        Write-Host "nothing here..."
    }

    $resolvedArray = $rArr
}

# Checking for equality on Severity count and creating a table.
if ($SevA_count) {
    $sevArrayCount += "
    <tr>
    <td>SevA</td>
    <td class='textcenter textred'>$SevA_count</td>
    </tr>
    "
}

if ($SevB_count) {
    $sevArrayCount += "
    <tr>
    <td>SevB</td>
    <td class='textcenter textred'>$SevB_count</td>
    </tr>
    "
}

if ($SevC_count) {
    $sevArrayCount += "
    <tr>
    <td>SevC</td>
    <td class='textcenter textred'>$SevC_count</td>
    </tr>
    "
}

if ($RegA_count) {
    $sevArrayCount += "
    <tr>
    <td>RegA</td>
    <td class='textcenter textred'>$RegA_count</td>
    </tr>
    "
}

if ($RegB_count) {
    $sevArrayCount += "
    <tr>
    <td>RegB</td>
    <td class='textcenter textred'>$RegB_count</td>
    </tr>
    "
}

if ($OpA_count) {
    $sevArrayCount += "
    <tr>
    <td>OpA</td>
    <td class='textcenter textred'>$OpA_count</td>
    </tr>
    "
}

if ($BCC_count) {
    $sevArrayCount += "
    <tr>
    <td>BCC</td>
    <td class='textcenter textred'>$BCC_count</td>
    </tr>
    "
}

$sevCountTable = "<table>
    <tr>
    <th class='Tb_heading' colspan='4'>Open Incidents broken by Severity type</th>
    </tr>
    <tr>
    <th>Severity</th>
    <th>Count</th>
    </tr>
    $sevArrayCount
    </table>
    <br>
    "

$sev_count += "<p>There were <span class='sevcount'>$res_highsev_count</span> Incidents resolved within 24 hours & there are currently <span class='sevcount'>$highsev_count</span> ongoing High Severity Incidents: <br><br>
$sevCountTable
</p>"

if ($resolvedArray) {
    $Resolved_SevTb = "<table class='tbWidth'>
    <tr>
    <th class='Tb_heading' colspan='4'>Resolved High Severity Incidents (24 Hours)</th>
    </tr>
    <tr>
    <th>Title</th>
    <th>Date</th>
    <th>INC</th>
    <th>Comm History</th>
    </tr>
    $resolvedArray
    </table>
    <br>"
}
else {
    #Clear-Variable $ALL_Sev_Table # this might clear the HTML Table variable if there are no sev PDs???
}#>


if ($openArray) {
    $ALL_Sev_Table = "<table class='tbWidth'>
    <tr>
    <th class='Tb_heading' colspan='5'>Open High Severity Incidents</th>
    </tr>
    <tr>
    <th>Title</th>
    <th>IC</th>
    <th>Date</th>
    <th>INC</th>
    <!--<th>Latest Comms</th>-->
    <th>Comm History</th>
    </tr>
    $openArray
    </table>
    <br>"
}
else {
    #Clear-Variable $ALL_Sev_Table # this might clear the HTML Table variable if there are no sev PDs???
}


# Begin HTML Construction
$HTMLmessage = @"
<html>
<head>
<style>
    tr:nth-child(even) {background-color: #f2f2f2;}
    TABLE{border: 1px solid black; border-collapse: collapse; font-size:12pt; font-family: Calibri;}
    .tbWidth{width: 100%; text-align: center;}
    TH{border: 1px solid black; background: #333333; padding: 5px; color: #ffffff;}
    TD{border: 1px solid black; padding: 5px;}
    .column_text_align{text-align: center !important;}
    .row{background: #000;}
    a {color: #004de0; text-decoration: none;}
    h1 {font-family: Calibri; padding: 0; margin: 0;}
    h3 {font-family: Calibri;}
    p {font-family: Calibri;}
    .null {background-color: #ffffff;}
    .Tb_heading {background-color: #333333; padding: 15px; font-size: 18px;}
    .Tb_heading2 {background-color: #333333; font-size: 18px;}
    .bg {background-color: #333333;}
    .sevcount {color: #ff0000; font-weight: bold; margin: 0; padding: 0;}
    .SevA {background-color: #ffe0e0 !important; text-align: center;}
    .SevB {background-color: #ffdec0 !important; text-align: center;}
    .SevC {background-color: #ffdb29 !important; text-align: center;}
    .RegA {background-color: #e5dded !important; text-align: center;}
    .RegB {background-color: #f4fdf0 !important; text-align: center;}
    .OpA {background-color: #afcce5 !important; text-align: center;}
    .textleft {text-align: left;}
    .textcenter {text-align: center !important;}
    .tdblack {background-color: #000000 !important;}
    .oldinc {background-color: #fff2f2 !important;}
    .textred {color: #ff0000 !important;}
</style>
<title>IC Handover Report</title>
</head>
<body>
<h1>High Severity Incident Report</h1>
$sev_count
$Resolved_SevTb
$ALL_Sev_Table
</body>    
</html>
"@
$HTMLmessage | Out-File -FilePath ".\index.html"


# Composing Email Message
$pass = ConvertTo-SecureString "#######" -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential ("ayush", $pass)

$From = "[EMAIL_ADDRESS]"
$To = "[EMAIL_ADDRESS]"
$Subject = "High Severity Report"
$SMTPServer = "[EMAIL_ADDRESS]"

try {
    Send-MailMessage -From $From -to $To -Subject $Subject -Body $HTMLmessage -BodyAsHtml -SmtpServer $SMTPServer -Credential $mycreds 
    Write-Host "Email sent!" -ForegroundColor Green
}
catch {
    Write-Host "Failed to send email..." -ForegroundColor Red
}
