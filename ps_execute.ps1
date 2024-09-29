# Define the GitHub repository and asset pattern (e.g., .exe file)
$repoOwner = "benlf1"
$repoName = "EqualizerAPOAuto"
$assetPattern = "*.exe"

# GitHub API URL for the latest release
$latestReleaseUrl = "https://api.github.com/repos/$repoOwner/$repoName/releases/latest"

# Get the latest release information from GitHub API
$releaseInfo = Invoke-RestMethod -Uri $latestReleaseUrl -Headers @{ "User-Agent" = "PowerShell Script" }

# Filter the assets to find the one matching the pattern (e.g., an .exe file)
$assetUrl = $releaseInfo.assets | Where-Object { $_.name -like $assetPattern } | Select-Object -ExpandProperty browser_download_url

# Define the path where the file will be downloaded
$tempFilePath = "$env:TEMP\$($assetUrl.Split('/')[-1])"

# Download the executable
Invoke-WebRequest -Uri $assetUrl -OutFile $tempFilePath

# Execute the downloaded executable as administrator
Start-Process -FilePath $tempFilePath -Verb RunAs -Wait

# Delete the executable after execution
Remove-Item -Path $tempFilePath -Force
