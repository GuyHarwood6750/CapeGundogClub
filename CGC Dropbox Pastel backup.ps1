   <#Backup CGC dropbox and Pastel data, scheduled task (same name).
   #>
   Compress-Archive -Path 'C:\Users\Guy\Dropbox\CGC' -DestinationPath "D:\CGC DROPBOX BACKUPS\CGC $(get-date -f yyyyMMdd-HHmmss).zip" -force
    Write-EventLog -LogName MyPowerShell -Source "CGC" -EntryType Information -EventId 10 -Message "Archive of CGC Dropbox data completed"

    Compress-Archive -Path 'C:\PASTEL19\CGC2021' -DestinationPath "D:\Pastel Backups\CGC\2021\CGC2021 $(get-date -f yyyyMMdd-HHmmss).zip" -force
    Write-EventLog -LogName MyPowerShell -Source "CGC" -EntryType Information -EventId 10 -Message "Archive of CGC2019 Pastel data completed"

    Compress-Archive -Path 'C:\Users\Guy\Dropbox\Cape Gundog Club' -DestinationPath "D:\CGC DROPBOX BACKUPS\CGCMagazine $(get-date -f yyyyMMdd-HHmmss).zip" -force