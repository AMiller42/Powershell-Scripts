$LogDays = Read-Host "Enter number of days to collect logs from - (Default 7)"
$LogPath = Read-Host "Enter folder path to save logs to - (Default $HOME\Documents)"
$LogSize = Read-Host "Enter number of entries to collect per log file - (Default 5000)"

if ( -not ($LogDays))
{
    $LogDays = 7
}

if ( -not ($LogPath))
{
    $LogPath = "$HOME\Documents"
}

if ( -not ($LogSize))
{
    $LogSize = 5000
}

#Connect to Office 365 
Connect-ExchangeOnline

#Collect Message Tracking Logs (These are broken into "pages" in Office 365 so we need to collect them all with a loop) 
$Messages = $null 
$Page = 1
Write-Host "Collecting messages for the last $LogDays days..."
do 
{ 
    $CurrMessages = Get-MessageTrace -StartDate (Get-Date).AddDays(-$LogDays) -EndDate (Get-Date)  -PageSize $LogSize  -Page $Page | Select Received,SenderAddress,FromIP,RecipientAddress,ToIP,Subject,Status,Size

    if ($CurrMessages -ne $null)
    {
        $CurrMessages | Export-Csv "$LogPath\PAGE-$Page.csv" -NoTypeInformation
        Write-Host "Collected Page $Page..."
    }
    $Page++ 
    $Messages += $CurrMessages 
} 
until ($CurrMessages -eq $null) 

Disconnect-ExchangeOnline -Confirm:$false