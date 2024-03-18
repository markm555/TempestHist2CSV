Function Get-Folder()
{
    #************************************************************************************
    #***                                                                              ***
    #***     Windows Form Function to select a folder where files will reside         ***
    #***                                                                              ***
    #************************************************************************************

    #[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    Add-Type -AssemblyName System.Windows.Forms

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.rootfolder = "MyComputer"
    #$foldername.ShowDialog()
    $result = $Foldername.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))

    if($result -eq "OK") {
        $folder += $foldername.SelectedPath
        return($folder)
    }
}

$path = get-folder

##############################################################################################
##               Change Values between these lines for your environment                     ##
##############################################################################################
$days = 1                   # Days back from today 1 = yesterday, 2 = the past two days
$deviceId = "<Device ID>"   # Your Device ID
$tokens = "<Access Token>"  # Your access Token
##############################################################################################

New-Item $path"\TempestHist.csv"

$url = "https://swd.weatherflow.com/swd/rest/observations/device/" + $deviceId + "?day_offset=" + $days + "&token=" + $tokens
$results = Invoke-WebRequest -uri $url -Method Get
$histraw = $results.Content |ConvertFrom-Json
$ob = $histraw.obs

$newRow = @()  
$object = New-Object -TypeName PSObject          
foreach($row in $histraw.obs)
{
   $object = New-Object PsObject
   $object | Add-Member -Name 'epoch' -MemberType Noteproperty -Value $row[0]
   $object | Add-Member -Name 'localTime' -MemberType Noteproperty -Value ((Get-Date -Date "01-01-1970") + ([System.TimeSpan]::FromSeconds($row[0]))).GetDateTimeFormats()[39]
   $object | Add-Member -Name 'windlull'    -MemberType Noteproperty -Value $row[1]
   $object | Add-Member -Name 'windavg'     -MemberType Noteproperty -Value $row[2]
   $object | Add-Member -Name 'windgust'    -MemberType Noteproperty -Value $row[3]
   $object | Add-Member -Name 'winddir'     -MemberType Noteproperty -Value $row[4]
   $object | Add-Member -Name 'windsample'  -MemberType Noteproperty -Value $row[5]
   $object | Add-Member -Name 'pressure'    -MemberType Noteproperty -Value $row[6]
   $object | Add-Member -Name 'airtemp'     -MemberType Noteproperty -Value $row[7]
   $object | Add-Member -Name 'relhumidity' -MemberType Noteproperty -Value $row[8]
   $object | Add-Member -Name 'Illuminance' -MemberType Noteproperty -Value $row[9]
   $object | Add-Member -Name 'uv'          -MemberType Noteproperty -Value $row[10]
   $object | Add-Member -Name 'solar'       -MemberType Noteproperty -Value $row[11]
   $object | Add-Member -Name 'rain'        -MemberType Noteproperty -Value $row[12]
   $object | Add-Member -Name 'preciptype'  -MemberType Noteproperty -Value $row[13]
   $object | Add-Member -Name 'strikedist'  -MemberType Noteproperty -Value $row[14]
   $object | Add-Member -Name 'strikecount' -MemberType Noteproperty -Value $row[15]
   $object | Add-Member -Name 'battery'     -MemberType Noteproperty -Value $row[16]
   $object | Add-Member -Name 'reptint'     -MemberType Noteproperty -Value $row[17]
   $object | Add-Member -Name 'dayrain'     -MemberType Noteproperty -Value $row[18]
   $object | Add-Member -Name 'ncrainacc'   -MemberType Noteproperty -Value $row[19]
   $object | Add-Member -Name 'dayrainacc'  -MemberType Noteproperty -Value $row[20]
   $object | Add-Member -Name 'precipatype' -MemberType Noteproperty -Value $row[21]
   $newRow += $object
}
   $newRow | Export-Csv -Path $path"\TempestHist.csv" -Append -NoTypeInformation
