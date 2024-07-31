$exclude = @("venv", "rpa_challenge.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "rpa_challenge.zip" -Force