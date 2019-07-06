#if($args.length -ne 1) {
#	write-host "Usage:cre-mmf.ps1
#	exit(1);
#}
#
$src_dir="o:\mp4"
$dest_dir="V:\movie-files"

$files=dir "$src_dir"|? { $_.extension -eq ".mp4" }|?  { !(Test-Path (join-path $dest_dir $_.name)) };

$files|
	foreach {
		$src=((join-path $_.directoryname $_.name));
		write "cp $src $dest_dir";
		cp -PassThru "$src" "$dest_dir";
	}

$files|
	foreach {
		$src=((join-path $dest_dir $_.name));
		dir "$src";
	}

