
Get-Content -Path C:\Windows\System32\LogFiles\HTTPERR\httperr1.log

#For example, the following command displays the Windows Server Update Services (WSUS) SoftwareDistribution log one page at a time.
Get-Content -Path 'C:\Program Files\Update Services\LogFiles\SoftwareDistribution.log' |Out-Host –Paging

#For instance, the following command displays the last 50 lines of the Deployment Image Servicing and Management (DISM) log file.
Get-Content -Path C:\Windows\Logs\DISM\dism.log -Tail 50

#In the next example, the command line displays the last five lines of the WindowsUpdate.log and waits for additional lines to display.
Get-Content -Path C:\Windows\WindowsUpdate.log -Tail 5 –Wait

#This searches all lines from the firewall log containing the word "Drop" and displays only the last 20 lines.
Select-String -Path C:\Windows\System32\LogFiles\Firewall\pfirewall.log ‑Pattern 'Drop' | Select-Object -Last 20

#For instance, the following command displays lines containing the word "error" or the word "warning" from the Windows Update agent log file.
Select-String -Path C:\Windows\WindowsUpdate.log -Pattern 'error','warning'

#The following command searches for lines with the word "err" preceded and followed by a space. It also displays the three lines before and after every match from the cluster log file.
Select-String C:\Windows\Cluster\Reports\Cluster.log -Pattern ' err ' ‑Context 3

#For instance, the following command line displays lines 45 to 75 from the netlogon.log file.
Get-Content C:\Windows\debug\netlogon.log |Select-Object -First 30 -Skip 45
