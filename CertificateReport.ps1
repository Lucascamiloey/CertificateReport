
##
## Changelog:
##

# V 0.5
# Agregar check de ips vivas. 
# log de error marca que pcs estan apagadas.
# ahora funciona de verdad.
#



## Inicializar los arrays
$serverlist=@()
$certlist=@()

## Trae la lista de servers que tengan "server" en el campo operatingSystem 
## Con esto filtramos los clusters PERO incluimos los nodos.
$Serverlist=(dsquery * -filter "(objectCategory=Computer)" -attr name operatingSystem -limit 10000)  | select @{label='ServerName';expression={(($_ -split ("   "))[0]).trim()}}, @{label='OperatingSystem';expression={(($_ -split ("   "))[1]  ).trim()}} | where {$_.operatingSystem -like "*server*"}

## Exporta la lista a un csv para que se pueda ejecutar el proceso de recoleccion de certs
$serverlist | export-csv -notypeinformation C:\temp\certs\servers.csv


## Proceso para traer los certs 

foreach ($server in $serverlist){

	$name=$server.servername
	
	#Chequea que el server este vivo.
	if (Test-Connection -ComputerName $name -Quiet -Count 1){
		# Descomentar la siguiente linea para debuggear psexec
		write-output "ejecutando psexec en $name"
		## -inputformat none hace que no se cuelgue despues de ejecutar
		psexec \\$name powershell.exe -inputformat none -Command "& {Get-ChildItem -Path cert:\LocalMachine\My | Select-Object Subject, Issuer, NotBefore, NotAfter, Thumbprint, SerialNumber |   Export-Csv -path c:\$name.csv -NoTypeInformation}"
		# Descomentar la siguiente linea para debuggear robocopy
		write-output "Robocopiando $name.csv"
		
		#Setear los parametros para el robocopy
		$source="\\$name\c$"
		$destination="c:\temp\certs\"
		$files="$name.csv"
		#Deprecamos $options porque no andaba bien
		#$options="/MOV /R:0 /W:3"
	
		#Hace el robocopy
		# | out-null es para que no escriba nada en la consola.
		robocopy $source $destination $files /MOV /R:0 /W:3 > $null
	
	}else{
		#Graba error en C:\temp\certs\logs\reportecerts-DIAMESAÑO-HORAMINUTO.log
		"OFFLINE: $name ">>"C:\temp\certs\logs\reportecerts-$(get-date -Format 'ddMMyy-hhmm').log"
		write-output "Error en $name"
	}
	
}

## Recorre la carpeta buscando CSVs y guarda las rutas en $list
$list=Get-ChildItem C:\Temp\certs\* -Include *.csv -Exclude servers.csv

## Recorre el array $list y recopila la informacion en $certlist
foreach ($file in $list){
	$servername=$file.name.trim(".csv")
	$certlist+=Import-Csv $file | select @{label='ServerName';expression={$servername}}, *
}

## Exporta el listado de certificados a CSV para que se pueda leer externamente
$certlist | export-csv -path C:\temp\certs\cert-list-full.csv 
