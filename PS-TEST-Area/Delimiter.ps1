#First Name,Last Name,Gender,Country,Age,Date,Id
[string]$string = "1,Dulce,Abril,Female,United States,32,15/10/2017,1562"


#decode
$lineNumber=$string.Split(",")[0]
$firstname=$string.Split(",")[1]
$lastname=$string.Split(",")[2]
$gender=$string.Split(",")[3]
$country=$string.Split(",")[4]
$age=$string.Split(",")[5]
$date=$string.Split(",")[6]
$Id=$string.Split(",")[7]

#$date.GetType().name
$newdate= [datetime]::parseexact($date,'dd/MM/yyyy',$null)
#$date.GetType().name
#$NewDate = Get-Date $newdate -Format "dd/MM/yyyy"
#$newdate=$newdate
$NewDate.GetType()

$age=[Int]$age
$id=[Int]$id

$lineNumber
$firstname
$lastname
$gender
$country
$age
$newdate
$Id





<#
$cnt = 1
foreach ($detailexplain in $string.Split(","))
{
    Write-Host "Element $cnt is $detailexplain"
    $cnt++
}
#>