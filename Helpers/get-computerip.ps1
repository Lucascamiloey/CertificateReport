##
##	Returns IPv4, IPv6 and status of a target computer
##

function get-computerip ($targetcomputer)
{
	$returnthis = New-Object �TypeName PSObject
	$returnthis | Add-Member �MemberType NoteProperty �Name Name -Value $targetcomputer
	$returnthis | Add-Member �MemberType NoteProperty �Name IPv4Address -Value ""
	$returnthis | Add-Member �MemberType NoteProperty �Name IPv6Address -Value ""
	$returnthis | Add-Member �MemberType NoteProperty �Name Online -Value $false
	
	$help=test-connection $targetcomputer -count 1 2>null
	if ($help){
		$returnthis.IPv4Address=$help.IPv4Address
		$returnthis.IPv6Address=$help.ProtocolAddress
		$returnthis.Online=$true
	} 
	
	return $returnthis
}