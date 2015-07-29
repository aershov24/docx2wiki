##convert.ps1##
param(
)

$Source = ".\Source"

Get-ChildItem $Source\ -filter *.docx | `
Foreach-Object{
    
$docx_file = ".\Source\"+$_

$homedir = (Get-Item -Path ".\" -Verbose).FullName
$zipdir = ".\Source\"+$_.BaseName
$wikidir = ".\wiki\"+$_.BaseName
$wikidirimage = ".\wiki\"+$_.BaseName+"\images"

New-Item $zipdir -type directory
New-Item $wikidir -type directory
New-Item $wikidirimage -type directory

$zip=$_ -replace ".docx", ".zip"
$zip_file = $zipdir +"\"+$zip

copy-item ".\Source\$_" $zip_file

cd $zipdir

$shell_app=new-object -com shell.application
$zip_arch = $shell_app.namespace((Get-Location).Path+"\$zip")
$destination = $shell_app.namespace((Get-Location).Path)
$destination.Copyhere($zip_arch.items())

cd $homedir

$wiki= $wikidir +"\" + $_ -replace ".docx", ".txt"
$wikiready= $wikidir+"\ready_"+ $_ -replace ".docx", ".txt"

$wikiimage = $zipdir +"\word\media"

.\pandoc -s $docx_file  -t mediawiki -o $wiki

$appendix = ([char[]]([char]'a'..[char]'z') + 0..9 | sort {get-random})[0..6] -join ''

Get-ChildItem -path $wikiimage | rename-item -newname {$_.Name -replace 'image', $appendix}
Get-Content $wiki | 
ForEach-Object { $_ -replace 'media/image', $appendix -replace '= =', ' ' -replace '= <br /> =', ' ' } | Set-Content ($wikiready)

copy-item "$wikiimage\*" $wikidirimage

}


##convert.ps1##