$Array = @()
$items = "1,2,3,4";
$itemlist = $items.split(",");
foreach ($item in $itemlist)

{
$Array += $item
}
$itemlist.GetType()
$Array.gettype()



#$Tracelog
#$errormessage
#$logtext.gettype()

#$LogLine = [System.Collections.ArrayList]::new()
#[void]$logline.Add($logtext.foreach{"Item [$PSItem]"})

#$LogLine.GetType()
#$logline

#$logcontent += $logtext.foreach{"Item [$PSItem]"}
#$logcontent

#foreach ($item in $logtext){[STRING]$logline=$item;$logline}


#foreach ($item in $logtext.Split("`n")){[STRING]$logline =$item}
#$logline=$logline.Split("`n")
#$logline