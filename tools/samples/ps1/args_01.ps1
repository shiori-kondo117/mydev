if ($args.length -ne 2) {
	$s=$myInvocation.MyCommand.name
	write-host "Usage:$s [arg1] [arg2]"
	exit 1;
}
