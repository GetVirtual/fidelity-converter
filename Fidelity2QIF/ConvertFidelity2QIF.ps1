# Turn on Debug
#$DebugPreference = "Continue"

# open file in script location
$scriptpath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Variables
[int]$i = 0

# Date of last sync to compare which lines are relevant, Format DD.MM.YYYY, MANUAL OVERRIDE
#[datetime]$lastsync = Get-Date ("11.01.2024")

# Import lastrun from file
[string]$lastrunfilecontent = Get-Content (Join-Path -Path $scriptpath -ChildPath "lastrun.txt")
[datetime]$lastsync = [datetime]::parseexact($lastrunfilecontent, 'dd.MM.yyyy', $null)

# import CSV
$CSVFile = Join-Path -Path $scriptpath -ChildPath "View open lots.csv"
$CSV = Import-Csv $CSVFile

# remove $QIFFILE if it exists
$QIFFILE = Join-Path -Path $scriptpath -ChildPath "Depot-Fidelity.qif"
if (Test-Path $QIFFILE) { Remove-Item $QIFFILE }

# remove $QIFFILE if it exists
$QIFFILE_VERRECHNUNG = Join-Path -Path $scriptpath -ChildPath "Verrechnung-Fidelity.qif"
if (Test-Path $QIFFILE_VERRECHNUNG) { Remove-Item $QIFFILE_VERRECHNUNG }


# REFERENCE FOR QIF 
# https://en.wikipedia.org/wiki/Quicken_Interchange_Format


$CSV | ForEach-Object {
    # New line
    Write-Debug $_

    try
    {
        # Check if line starts with a date
        [datetime]$date = Get-Date $_.('Date acquired')
    }
    catch
    {
        # if line starts with an empty string break loop
        return
    }

    # Variables
    [string]$QIF_ENTRY = ""
    [string]$QIF_VERRECHNUNG_ENTRY = ""
    [double]$costbasis = 0.00
    [string]$qifdate = ""


    # only proceed if $date is older then lastsync
    [datetime]$objectdate = Get-Date $_.('Date acquired')
    if ($objectdate -lt $lastsync) { 
        Write-Debug "Skipping - Item from $objectdate, because it is older then $lastsync"
        return # use return instead of continue, because continue would skip the whole loop
    }
    else {
        Write-Debug "Continue - Item from $objectdate, because it is newer then $lastsync"
    }
    
    # Reformat date
    $qifdate = (Get-Date $_.('Date acquired') -Format "M.d.yyyy")
    
    
    ###################################################################
    # Depot
    ###################################################################
        
    $QIF_ENTRY = "!Option:MDY"
    $QIF_ENTRY += "`n!Type:Invst"

    # Date
    # Convert Mar-29-2018 to 3.29.2018
    $QIF_ENTRY += "`nD" + $qifdate
    $QIF_ENTRY += "`nV" + $qifdate

    #$_.('Date acquired')

    $QIF_ENTRY += "`nNKauf"
    $QIF_ENTRY += "`nFEUR"
    $QIF_ENTRY += "`nG1.000000"

    # Source of the share - Comment Field
    if ($_.('Share source') -eq "DO") 
        { $QIF_ENTRY += "`nMVesting" }
    elseif ($_.('Share source') -eq "SP") 
        { $QIF_ENTRY += "`nMESPP" }

    $QIF_ENTRY += "`nYMICROSOFT CORP. - REGISTERED SHARES DL -,00000625 Xetra"
    $QIF_ENTRY += "`n~870747"
    $QIF_ENTRY += "`n@US5949181045"

    $QIF_ENTRY += "`n&1"
    
    # Value per share
    $QIF_ENTRY += "`nI" + $_.('Cost basis/share')

    # Quantity
    $QIF_ENTRY += "`nQ" + $_.('Quantity')

    # Calculate Cost basis (Quantity * Cost basis/share)
    $costbasis = [double]$_.('Quantity') * [double]$_.('Cost basis/share')
    # round to 2 decimals
    $costbasis = [math]::Round($costbasis, 2)

    # Total
    $QIF_ENTRY += "`nU" + $costbasis


    $QIF_ENTRY += "`nO0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00|0.00"
    
    # Verrechnungskonto
    $QIF_ENTRY += "`nL|[Fidelity Verrechnung]"
    
    # Total again?
    $QIF_ENTRY += "`n$" + $costbasis

    $QIF_ENTRY += "`nB0.00|0.00|0.00"
    $QIF_ENTRY += "`n^"
    
    # Write to File
    $QIF_ENTRY | Out-File $QIFFILE -Encoding utf8 -Append

    ###################################################################
    # Verrechnungskonto
    ###################################################################

    
    $QIF_VERRECHNUNG_ENTRY = "!Option:MDY"
    $QIF_VERRECHNUNG_ENTRY += "`n!Type:Oth T"

    # Date
    $QIF_VERRECHNUNG_ENTRY += "`nD" + $qifdate

    # Value
    $QIF_VERRECHNUNG_ENTRY += "`nU" + $costbasis
    $QIF_VERRECHNUNG_ENTRY += "`nT" + $costbasis


    $QIF_VERRECHNUNG_ENTRY += "`nCX"

    # Source of the share defines Comment & Category
    if ($_.('Share source') -eq "DO") { 
        $QIF_VERRECHNUNG_ENTRY += "`nMVesting"
        $QIF_VERRECHNUNG_ENTRY += "`nLLohn-Gehalt:Vesting"

    }
    elseif ($_.('Share source') -eq "SP") { 
        $QIF_VERRECHNUNG_ENTRY += "`nMESPP"
        $QIF_VERRECHNUNG_ENTRY += "`nLLohn-Gehalt:ESPP"
    }

    # Empfänger
    $QIF_VERRECHNUNG_ENTRY += "`nPAktienzuweisung"

    $QIF_VERRECHNUNG_ENTRY += "`n^"

    # Write to File
    $QIF_VERRECHNUNG_ENTRY | Out-File $QIFFILE_VERRECHNUNG -Encoding utf8 -Append

    $i++

}

Write-Host "Written $i items to $QIFFILE"
Write-Host "Written $i items to $QIFFILE_VERRECHNUNG"

# Export today to file
if ($i -gt 0) {
    [string]$today = Get-Date -Format "dd.MM.yyyy"
    $today | out-file (Join-Path -Path $scriptpath -ChildPath "lastrun.txt")
}






