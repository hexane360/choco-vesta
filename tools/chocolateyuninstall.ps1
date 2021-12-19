$ErrorActionPreference = 'Stop'

$out_dir = Split-Path -Parent $(Split-Path -Parent $MyInvocation.MyCommand.Definition)
$bin_path = Join-Path $out_dir 'vesta.exe'

# remove shim
Uninstall-BinFile 'vesta.exe' $bin_path

# and remove start menu shortcut
foreach ($start_kind in @('StartMenu', 'CommonStartMenu')) {
	$start_dir = Join-Path $([Environment]::GetFolderPath($start_kind)) 'Programs\VESTA'
	if ([System.IO.Directory]::Exists($start_dir)) {
		[System.IO.Directory]::Delete($start_dir, $true)
		$user = if ($start_kind -eq 'StartMenu') { 'user' } else { 'system' }
		echo "Removed $user start menu shortcut."
	}
}

# and remove file associations
$ftype = "VESTA"
if (Test-Path "Registry::HKCR\$ftype") {
	Remove-Item -Path "Registry::HKCR\$ftype" -Recurse
	echo "Removed file assocations."
}