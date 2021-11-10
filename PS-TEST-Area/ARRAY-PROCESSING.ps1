$Array = @()
$items = "1,2,3,4";
$itemlist = $items.split(",");
foreach ($item in $itemlist)

{
$Array += $item
}

$Array