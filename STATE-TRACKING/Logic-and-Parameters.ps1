$displayname ="Alec Guiness"
$displayname=$displayname.Split(" ")
$initial=$displayname[0].Substring(0,1)
$lastname=$displayname[1]
$SamAccountName ="$initial$lastname"

#$SamAccountName

#create runbook tracking guid
$RBGuid=[guid]::NewGuid()
$RBGuid

<#use SCORCHPersistantDB
SELECT Displayname,
ActivityName,
Activitystatus
from StateTrackin01
#>

$SQLdata ="Alec Guiness";"Create AD Account";"failed"
#$sqldata=$null

If($SQLData){
$data=$SQLDATA.Split(";")

$ActivityName=$data[0]
$activityStatus=$data[1]
$DisplayName=$data[2]

$ActivityName
$activityStatus
$DisplayName
}else {
  $action="Create Account"  
}

# parse content and decide action