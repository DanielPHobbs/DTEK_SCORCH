#https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/write-eventlog?view=powershell-5.1


Write-EventLog -LogName "Application" -Source "MyApp" -EventID 3001 -EntryType Information -Message "MyApp added a user-requested feature to the display." -Category 1 -Ra

