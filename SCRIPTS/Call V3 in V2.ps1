#prepare data to pass
$databusvar = "\`d.T.~Ed/{2B4BE08D-BD2D-4ACD-856C-83764177F88B}.databusvar\`d.T.~Ed/"
$someotherthing = "someother"
$inobj = new-object pscustomobject -property @{
    databusvar = $databusvar
    other=$someotherthing 
 }



#call powershell V3
$theresults = $inobj | PowerShell {
    
    #use the special $input variable, just get the first item in case multiple
    #objects were piped in. (which weren't in this case)
    $inobject = $input | select -first 1
    #return results 
    
    
    
    
    new-object pscustomobject -property @{        
        version = " from Version $($PSVersionTable.psversion.tostring())"
        databusuppercase = $inobject.databusvar.toupper()
        hellorunbook = "hello $($inobject.other)"
       }
 }
#take the results from property and put them in variables for
#the invoke.net script activity to pick up and publish on the databus
$theversion = $theresults.version
$other = $theresults.hellorunbook
$databusvar = $theresults.databusuppercase