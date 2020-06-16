param(
$collectionURL,
$projectName,
$processName
)

$scriptFolder = Split-Path -Path $MyInvocation.MyCommand.Path

$logfile= $scriptFolder + "\witdlog.txt"

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

$witdfile = ""

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

if($processName -eq "agile") {
  $witdfile = Get-Content ($scriptFolder + "\AgileWITD.txt")
}

if($processName -eq "scrum"){
  $witdfile = Get-Content ($scriptFolder + "\SCRUMWITD.txt")
}

$executionStartTime = Get-Date

"## Renomage des éléments de travail dans Azure DevOps Server @ $executionStartTime ##"| Out-File $logfile -Append
"Nom du projet : $projetName, Processus utilisé : $processName" | Out-File $logfile -Append

foreach ($WIID in $witdfile)
    {
        $params = $WIID.Split(",")
        $fname = $params[0]
        $ename = $params[1]

         Invoke-Expression "& $WitAdminExe renamewitd /collection:$CollectionUrl /p:`"$projectName`" /n:`"$fname`" /new:`"$ename`" /noprompt"

        "$fname renomé en $ename" | Out-File $logfile -Append
        $WICount = $WICount + 1
    }

$ExecutionEndTime = Get-Date
"## Fin de l'opération @ $executionEndTime ##"| Out-File $logfile -Append