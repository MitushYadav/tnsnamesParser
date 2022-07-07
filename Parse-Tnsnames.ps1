function Parse-Tnsnames {

[CmdletBinding()]
param(
[string]
$PathToTnsnamesFile,
[string]
$RegexFolder,
[string]
$RegexFilePrefix,
[string]
$OutputFolder
)

$headerhashtable1 = [ordered]@{
    1 = 'NetServiceName'
    2 = 'Protocol'
    3 = 'Host'
    4 = 'Port'
    5 = 'Server'
    6 = 'ServiceName'
}

$headerhashtable2 = [ordered]@{
    1 = 'NetServiceName'
    2 = 'Protocol'
    3 = 'Host'
    4 = 'Port'
    5 = 'ServiceName'
}

$headerhashtable3 = [ordered]@{
    1 = 'NetServiceName'
    2 = 'Protocol'
    3 = 'Host'
    4 = 'Port'
    5 = 'ServiceName'
}

$headerhashtable4 = [ordered]@{
    1 = 'NetServiceName'
    2 = 'Protocol'
    3 = 'Host'
    4 = 'Port'
    5 = 'ServiceName'
}

$headerhashtable5 = [ordered]@{
    1 = 'NetServiceName'
    2 = 'Protocol1'
    3 = 'Host1'
    4 = 'Port1'
    5 = 'Protocol2'
    6 = 'Host2'
    7 = 'Port2'
    8 = 'Failover'
    9  = 'Server'
    10 = 'ServiceName'
}

$headerhashtable6 = [ordered]@{
    1 = 'NetServiceName'
    2 = 'Protocol'
    3 = 'Host'
    4 = 'Port'
    5 = 'ServiceName'
    6 = 'InstanceName'
}

$headerhashtable7 = [ordered]@{
    1 = 'NetServiceName'
    2 = 'RetryCount'
    3 = 'RetryDelay'
    4 = 'Protocol'
    5 = 'Port'
    6 = 'Host'
    7 = 'ServiceName'
    8 = 'SSLServerCertDN'
}

$headerhashtable8 = [ordered]@{
    1 = 'NetServiceName'
    2 = 'Protocol'
    3 = 'Host'
    4 = 'Port'
    5 = 'Server'
    6 = 'ServiceName'
}

$headerhashtable9 = [ordered]@{
    1 = 'NetServiceName'
    2 = 'Protocol'
    3 = 'Host'
    4 = 'Port'
    5 = 'Server'
    6 = 'ServiceName'
    7 = 'InstanceName'
}

$headerhashtable10 = [ordered]@{
    1 = 'NetServiceName'
    2 = 'Protocol'
    3 = 'Host'
    4 = 'Port'
    5 = 'Server'
    6 = 'UR'
    7 = 'ServiceName'
}

$tnsnames = Get-Content -Path $PathToTnsnamesFile -raw
$regexOptions = [Text.RegularExpressions.RegexOptions]::Multiline + [Text.RegularExpressions.RegexOptions]::IgnoreCase

$sum = 0
foreach($i in 1..10) {
$filePath = Join-Path -Path $RegexFolder -ChildPath "$RegexFilePrefix_$i.txt"
$regexPattern = Get-Content -Path $filePath -raw
$regMatches = [regex]::Matches($tnsnames, $regexPattern, $regexOptions)
Write-Host "Number of matches for number $i : " + $($regMatches | Measure-Object).Count
$sum += $($regMatches | Measure-Object).Count
$data = [System.Collections.Generic.List[Object]]@()
ForEach($rm in $regMatches) {
    $obj = New-Object PSObject
    $ht = Get-Variable -Name $('headerhashtable'+$i) -ValueOnly
    $ht.GetEnumerator() | ForEach-Object { Add-Member -InputObject $obj -MemberType NoteProperty -Name $($PsItem.Value) -Value $( $rm.Groups | where Name -eq $PsItem.Name | select -ExpandProperty Value )  }        
    $data.Add($obj) | out-null
    $sampleFormat = $rm.Groups | Where-Object Name -eq 0 | Select-Object -ExpandProperty Value
}

$data | Export-Excel -Path "$OutputFolder\parsed_$i.xlsx" -WorksheetName Entries
$sampleFormat | Export-Excel -Path "$OutputFolder\parsed_$i.xlsx" -WorksheetName SampleFormat

}
Write-Host "Sum: $sum"
}