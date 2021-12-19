$ErrorActionPreference = 'Stop'

$version = '3.5.7'
$checksum32 = '35C0E7483781398163A5E1472898E4243426691977CA6AEE7B4096A41FC10024'
$checksum64 = '7488ACE5AE3A960582E64717A5FBE642F6B4A1E8A715FBD9985F2B836A5E1760'
$urlBase = 'https://jp-minerals.org/vesta/archives/'

$exts = @('.vesta', '.cif', '.cmt', '.cssr', '.csd', '.icsd', '.pdb', '.xyz', '.xtl')

$out_dir = Split-Path -Parent $(Split-Path -Parent $MyInvocation.MyCommand.Definition)
$bin_path = Join-Path $out_dir 'vesta.exe'

$pp = Get-PackageParameters
$is_admin = Test-ProcessAdminRights

# download zip
$args = @{
	package        = $packageName
	unzipLocation  = $out_dir
	url64bit       = "$urlBase/$version/VESTA-win64.zip"
	checksum64     = $checksum64
	checksumType64 = 'sha256'
	url            = "$urlBase/$version/VESTA.zip"
	checksum32     = $checksum32
	checksumType32 = 'sha256'
}
Install-ChocolateyZipPackage @args

# flatten archive output
$dirs = Get-ChildItem $out_dir -Directory -Include 'VESTA*'
foreach ($dir in $dirs) {
	foreach ($item in $(Get-ChildItem $dir)) {
		$dest = Join-Path $out_dir $item.Name
		# ensure destination is empty before moving
		if (Test-Path $dest) {
			Remove-Item $dest -Recurse -Force
		}
		$item.MoveTo($dest)
	}
	$dir.Delete($false)
}

# hide executables
$exes = Get-ChildItem $out_dir -Include '*.exe' -Recurse
foreach ($exe in $exes) {
	New-Item "$exe.ignore" -Type file -Force | Out-Null
}
# install binary shim
Install-BinFile 'vesta' $bin_path -UseStart

# install start shortcut
if (!$pp['NoStart']) {
	$start_kind = if ($pp['UserStart'] -or !$is_admin) { 'StartMenu' } else { 'CommonStartMenu' }
	$start_dir = Join-Path $([Environment]::GetFolderPath($start_kind)) 'Programs'
	$shortcut_path = $(Join-Path $start_dir 'VESTA\VESTA.lnk')

	$args = @{
		shortcutFilePath = $shortcut_path
		targetPath = $bin_path
		workingDirectory = $out_dir
	}
	Install-ChocolateyShortcut @args

	$user = if ($start_kind -eq 'StartMenu') { 'user' } else { 'system' }
	echo "Added $user start menu shortcut."
}

# add file associations
if ($pp['FileAssoc']) {
	if ($is_admin) {
		$ftype = "VESTA"
		$desc = "VESTA file"
		$elevated_cmds = foreach ($ext in $exts) {
			"assoc $ext=$ftype"
		}
		$elevated_cmds += @(
			"ftype $ftype=`"$bin_path`" `"%1`" %*",
			"reg add HKCR\$ftype /ve /d `"$desc`""
		)
		$args = @("/c", @($elevated_cmds -join " && "))
		Start-ChocolateyProcessAsAdmin $args $env:COMSPEC
		echo "Added $($exts.Length) file assocation(s)."
	} else {
		Write-Warning "Not an administrator. Not adding file associations."
	}
}