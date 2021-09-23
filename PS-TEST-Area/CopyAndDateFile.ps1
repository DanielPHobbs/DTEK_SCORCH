
Try{

$timestamp = Get-Date -Format o | ForEach-Object { $_ -replace ":", "." }
$daytime=$(get-date -f MM-dd-yyyy_HH_mm_ss) 


$fileSoure="\\dteksan1\FileSource\TEXT\stock.txt"
$fileDestinationDir="\\dteksan1\FileDrop\"

Remove-Item -Path "$fileDestinationDir*.*" -Force


$Day=$(get-date -f dd-MM-yyyy) 
$newfileName="stock-$day.txt"

$fileDestination="$fileDestinationDir$newfileName"

Copy-Item "$fileSoure" -Destination "$fileDestination" -Force

}
Catch{

    $ErrorMessage = $_.Exception.Message
}

