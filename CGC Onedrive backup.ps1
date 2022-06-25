   <#Backup CGC dropbox and Pastel data, scheduled task (same name).
   #>
   Compress-Archive -Path 'C:\Users\Guy\OneDrive\Cape Gundog Club' -DestinationPath "D:\CGC OneDrive BACKUPS\CGC $(get-date -f yyyyMMdd-HHmmss).zip" -force
    Write-EventLog -LogName MyPowerShell -Source "CGC" -EntryType Information -EventId 10 -Message "Backup of CGC OneDrive data completed"