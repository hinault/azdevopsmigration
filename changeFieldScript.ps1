param($collectionURL
)

$VSDirectories = @()
$VSDirectories += "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"
$VSDirectories += "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Professional\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"
$VSDirectories += "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"
$VSDirectories += "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\TeamExplorer\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"
$VSDirectories += "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Community\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"
$VSDirectories += "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Professional\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"
$VSDirectories += "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"
$VSDirectories += "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\TeamExplorer\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"
$VSDirectories += "${env:ProgramFiles(x86)}\Microsoft Visual Studio 14.0\Common7\IDE"
$VSDirectories += "${env:ProgramFiles(x86)}\Microsoft Visual Studio 12.0\Common7\IDE"
$VSDirectories += "${env:ProgramFiles(x86)}\Microsoft Visual Studio 11.0\Common7\IDE"
$VSDirectories += "${env:ProgramFiles(x86)}\Microsoft Visual Studio 10.0\Common7\IDE"

$WitAdminExe = "witadmin.exe"

$scriptFolder = Split-Path -Path $MyInvocation.MyCommand.Path

$WI_List =  Get-Content ($scriptFolder + "\ChangeField.txt")
$logfile = $scriptFolder + "\log.txt"
$ExecutionStartTime = Get-Date
$WICount = 0


if(-not (Get-Command $WitAdminExe -ErrorAction SilentlyContinue)) {
  Write-Host -Verbose "Unable to find witadmin.exe on your path. Attempting VS install directories"
  foreach($vsDir in $VSDirectories) {
    $WitAdminExe = Join-Path $vsDir "witadmin.exe"
    Write-Host -Verbose "Testing for $WitAdminExe"
    if(Test-Path $WitAdminExe) {
      break
    }
  }
}

if(-not (Test-Path $WitAdminExe)) {
  throw "Unable to find the witadmin.exe in your path or in any VS install."
}

# Format witadmin exe with quotes for the invoke-expression to like
$WitAdminExe = "'$WitAdminExe'"


"## Renomage des champs dans Azure DevOps Server @ $ExecutionStartTime ##"| Out-File $logfile -Append
"URL de la collection: $CollectionURL" | Out-File $logfile -Append

 CD $WitAdminLocation

foreach ($WIID in $WI_List)
    {
        $params = $WIID.Split(",")
        $champ = $params[0]
        $val = $params[1]

        Invoke-Expression "& $WitAdminExe changefield /collection:$CollectionUrl /n:`"$champ`" /name:`"$val`" /noprompt"
     
        "$champ renomé" | Out-File $logfile -Append
        $WICount = $WICount + 1
    }

$ExecutionEndTime = Get-Date
"## Fin de l'opération @ $ExecutionEndTime ##"| Out-File $logfile -Append

"Nombre totale d'élément renommés: $WICount"   | Out-File $logfile -Append

##End of script##