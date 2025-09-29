# Turn on Debug
$DebugPreference = "Continue"

# open file in script location
$scriptpath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Variables
[int]$i = 0

# Date of last sync to compare which lines are relevant, Format DD.MM.YYYY, MANUAL OVERRIDE
[datetime]$lastsync = Get-Date ("01.06.1990")

# Import lastrun from file
#[string]$lastrunfilecontent = Get-Content (Join-Path -Path $scriptpath -ChildPath "lastrun.txt")
#[datetime]$lastsync = [datetime]::parseexact($lastrunfilecontent, 'dd.MM.yyyy', $null)

# import CSV
$CSVFile = Join-Path -Path $scriptpath -ChildPath "View open lots.csv"
$CSV = Import-Csv $CSVFile

# remove $QIFFILE if it exists
$FinanzflussCSVFILE = Join-Path -Path $scriptpath -ChildPath "Finanzfluss-Fidelity-Import.csv"
if (Test-Path $FinanzflussCSVFILE) { Remove-Item $FinanzflussCSVFILE }

# REFERENCE FOR QIF 
# https://en.wikipedia.org/wiki/Quicken_Interchange_Format

# Header
# Datum;ISIN;Name;Typ;Transaktion;Preis;Anzahl;Gebühren;Steuern;Währung;Wechselkurs
$FinanzflussCSVFILE_ENTRY = "Datum;ISIN;Name;Typ;Transaktion;Preis;Anzahl;Gebühren;Steuern;Währung;Wechselkurs"
$FinanzflussCSVFILE_ENTRY | Out-File $FinanzflussCSVFILE -Encoding utf8


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
    [string]$FinanzflussCSV_ENTRY = ""
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
    # Datum
    # ISIN
    # Name
    # Typ
    # Transaktion
    # Preis
    # Anzahl
    # Gebühren
    # Steuern
    # Währung
    # Wechselkurs
    ###################################################################
    
    # date (dd.MM.yyyy oder yyyy-MM-dd Beispiel: 2021-11-07)
    $FinanzflussCSV_date = (Get-Date $_.('Date acquired') -Format "dd.MM.yyyy")      # Reformat date
    $FinanzflussCSV_ENTRY = $FinanzflussCSV_date + ";"    

    # Microsoft Stock ISIN
    $FinanzflussCSV_ENTRY += "US5949181045;"
    
    # Name
    $FinanzflussCSV_ENTRY += "Microsoft;"
    
    # Typ
    $FinanzflussCSV_ENTRY += "Aktie;"

    # Transaktion
    $FinanzflussCSV_ENTRY += "Kauf;"

    # Preis
    $FinanzflussCSV_ENTRY += $_.('Cost basis/share').Replace(".", ",") + ";"

    # Anzahl
    $FinanzflussCSV_ENTRY += $_.('Quantity').Replace(".", ",") + ";"

    # Gebühren
    $FinanzflussCSV_ENTRY += "0;"

    # Steuern
    $FinanzflussCSV_ENTRY += "0;"

    # Währung
    $FinanzflussCSV_ENTRY += "USD;"

    # Wechselkurs
    #$exchangeRate = Invoke-RestMethod -Uri "https://api.exchangerate.host/convert?from=USD&to=EUR&date=$FinanzflussCSV_date" | Select-Object -ExpandProperty info | Select-Object -ExpandProperty rate
    $exchangeRate = "1,00"
    $FinanzflussCSV_ENTRY += $exchangeRate
        
    
    
<#     # Source of the share - Comment Field
    if ($_.('Share source') -eq "DO") 
        { $QIF_ENTRY += "`nMVesting" }
    elseif ($_.('Share source') -eq "SP") 
        { $QIF_ENTRY += "`nMESPP" }

#>
   
    
    # Write to File
    $FinanzflussCSV_ENTRY | Out-File $FinanzflussCSVFILE -Encoding utf8 -Append

    $i++

}

Write-Host "Written $i items to $FinanzflussCSVFILE"

# Export today to lastrun file
if ($i -gt 0) {
    [string]$today = Get-Date -Format "dd.MM.yyyy"
    $today | out-file (Join-Path -Path $scriptpath -ChildPath "lastrun.txt")
}







