###----------------------------------------------------

# Mode   [Sleep] =1 --- [Sleep till condition=True] =2

# Duration  in Seconds

# Condition = True/False

#-------------------------------------------------------

[bool]$condition=$true
[int]$mode =1
[int]$duration = 60

If($Mode -eq 1){

Start-Sleep -Seconds $duration

}else{

# ----------- Pause for 60 seconds per loop -----------
Do {
    # Do stuff
    # Sleep 5 seconds
    Start-Sleep -s $duration
}
while ($condition -eq $true)

}