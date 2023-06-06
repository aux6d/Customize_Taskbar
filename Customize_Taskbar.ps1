$Addfiles = (
	"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.exe",
	"C:\Program Files (x86)\Microsoft Office\Office16\excel.exe",
	"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.exe",
	"C:\Program Files (x86)\Microsoft Office\Office16\ONENOTE.exe",
	"C:\Program Files (x86)\Internet Explorer\iexplore.exe",
	"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
)
$Removefiles = ("C:\Windows\system32\ServerManager.exe",
	"C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe"
)

function Get-FileVerbs{
	Param(
		[System.IO.FileInfo]$file
	)
	$file |Format-List *
	$shell=new-object -ComObject Shell.Application
	$ns=$shell.NameSpace($file.DirectoryName)
	$ns.ParseName($file.Name).Verbs()
}

foreach ($file in $Addfiles)
{
"ADD $file"
$verbs = Get-FileVerbs $file
$verbs |Where-Object{$_.Name -eq 'Épingler à la &barre des tâches'} | ForEach-Object{$_.Doit()}
}

foreach ($file in $Removefiles)
{
"DEL $file"
$verbs = Get-FileVerbs $file
$verbs |Where-Object{$_.Name -eq 'Détacher de la &barre des tâches'} | ForEach-Object{$_.Doit()}
}
