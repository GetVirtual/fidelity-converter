# Turn on Debug
#$DebugPreference = "Continue"

# open file in script location
$scriptpath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Variables
[int]$i = 0

# Date of last sync to compare which lines are relevant, Format DD.MM.YYYY, MANUAL OVERRIDE
#[datetime]$lastsync = Get-Date ("01.01.1900")

# Import lastrun from file
[string]$lastrunfilecontent = Get-Content (Join-Path -Path $scriptpath -ChildPath "lastrun.txt")
[datetime]$lastsync = [datetime]::parseexact($lastrunfilecontent, 'dd.MM.yyyy', $null)

# import CSV
$CSVFile = Join-Path -Path $scriptpath -ChildPath "..\View open lots.csv"
$CSV = Import-Csv $CSVFile

# remove $QIFFILE if it exists
$ParqetCSVFILE = Join-Path -Path $scriptpath -ChildPath "Parqet-Fidelity-Import.csv"
if (Test-Path $ParqetCSVFILE) { Remove-Item $ParqetCSVFILE }

# REFERENCE FOR QIF 
# https://en.wikipedia.org/wiki/Quicken_Interchange_Format

# Header
$ParqetCSV_ENTRY = "date;currency;fee;tax;identifier;price;shares;type"
$ParqetCSV_ENTRY | Out-File $ParqetCSVFILE -Encoding utf8



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
    [string]$ParqetCSV_ENTRY = ""
    #[double]$costbasis = 0.00 #?
    

    # only proceed if $date is older then lastsync
    [datetime]$objectdate = Get-Date $_.('Date acquired')
    if ($objectdate -lt $lastsync) { 
        Write-Debug "Skipping - Item from $objectdate, because it is older then $lastsync"
        return # use return instead of continue, because continue would skip the whole loop
    }
    else {
        Write-Debug "Continue - Item from $objectdate, because it is newer then $lastsync"
    }
    
    ###################################################################
    # CSV File (Semi Colon as Separator!!!)
    # Header
    # date (dd.MM.yyyy oder yyyy-MM-dd Beispiel: 2021-11-07)
    # currency (EUR, USD)
    # fee
    # tax
    # fxrate (optional, umrechnungskurs für den Tag der Transaktion)
    # identifier (ISIN)
    # price (price / share - Positive Dezimalzahl ohne Tausendertrennzeichen Beispiel: 1,00)
    # shares (anzahl shares)
    # amount (summe der transaktion - wirklich mandatory?)
    # type (Buy)
    ###################################################################
    
    # date (dd.MM.yyyy oder yyyy-MM-dd Beispiel: 2021-11-07)
    $ParqetCSV_date = (Get-Date $_.('Date acquired') -Format "dd.MM.yyyy")      # Reformat date
    $ParqetCSV_ENTRY = $ParqetCSV_date + ";"    

    # currency (EUR, USD)
    $ParqetCSV_ENTRY += "USD;"
    
    # fee
    $ParqetCSV_ENTRY += "0;"
    
    # tax
    $ParqetCSV_ENTRY += "0;"
    
    # fxrate (optional, umrechnungskurs für den Tag der Transaktion) --> skipped
    
    # identifier (ISIN)
    $ParqetCSV_ENTRY += "US5949181045;"
    
    # price (price / share - Positive Dezimalzahl ohne Tausendertrennzeichen Beispiel: 1,00)
    # replace . with ,
    $ParqetCSV_ENTRY += $_.('Cost basis/share').Replace(".", ",") + ";"
    #$ParqetCSV_ENTRY += $_.('Cost basis/share') + ";"
    
    # shares (anzahl shares)
    # replace . with ,
    $ParqetCSV_ENTRY += $_.('Quantity').Replace(".", ",") + ";"
    #$ParqetCSV_ENTRY += $_.('Quantity') + ";"
    
    # amount (summe der transaktion - wirklich mandatory?)
    
    # type (Buy)
    $ParqetCSV_ENTRY += "Buy"


    
<#     # Source of the share - Comment Field
    if ($_.('Share source') -eq "DO") 
        { $QIF_ENTRY += "`nMVesting" }
    elseif ($_.('Share source') -eq "SP") 
        { $QIF_ENTRY += "`nMESPP" }

#>
   
    
    # Write to File
    $ParqetCSV_ENTRY | Out-File $ParqetCSVFILE -Encoding utf8 -Append

    $i++

}

Write-Host "Written $i items to $ParqetCSVFILE"

# Export today to lastrun file
if ($i -gt 0) {
    [string]$today = Get-Date -Format "dd.MM.yyyy"
    $today | out-file (Join-Path -Path $scriptpath -ChildPath "lastrun.txt")
}







