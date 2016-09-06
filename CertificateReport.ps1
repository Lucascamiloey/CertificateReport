
##
## Changelog:
##

# V 0.60
# Log de error en un mismo archivo por corrida
# Agregado el chequeo verbose con show-progressbar
# Ahora solo trae columnas relevantes
#


# Agregando funciones
function show-progressbar([int]$actual,[int]$completo, $estado, $actividad)
{
	$porcentaje=($actual/$completo)*100
	if (!$estado){
		$estado="Buscando datos $actual de $completo"
	}
	if (!$actividad){
		$actividad="Obteniendo Resultados"
	}
	Write-Progress -Activity $actividad -status $estado -percentComplete $porcentaje
}




## Inicializar los arrays
$serverlist=@()
$certlist=@()
$numeroactual=0
$fecha=get-date -Format 'ddMMyy-hhmm'

## Trae la lista de servers que tengan "server" en el campo operatingSystem 
## Con esto filtramos los clusters PERO incluimos los nodos.
$Serverlist=(dsquery * -filter "(objectCategory=Computer)" -attr name operatingSystem -limit 10000)  | select @{label='ServerName';expression={(($_ -split ("   "))[0]).trim()}}, @{label='OperatingSystem';expression={(($_ -split ("   "))[1]  ).trim()}} | where {$_.operatingSystem -like "*server*"}

## Exporta la lista a un csv para que se pueda ejecutar el proceso de recoleccion de certs
$serverlist | export-csv -notypeinformation C:\temp\certs\servers.csv
$total=$serverlist.count

## Proceso para traer los certs 

foreach ($server in $serverlist){
	#sumando iteracion
	$numeroactual=$numeroactual+1
	$name=$server.servername
	
	#muestra progreso
	$estado="Revisando $numeroactual de $total - $name"
	$actividad="Revisando"
	show-progressbar $numeroactual $total $estado $actividad
	
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
		
		#muestra progreso
		$estado="Revisando $numeroactual de $total - $name"
		$actividad="Revisando"
		show-progressbar $numeroactual $total $estado $actividad
	
	}else{
		#Graba error en C:\temp\certs\logs\reportecerts-DIAMESAÑO-HORAMINUTO.log
		"OFFLINE: $name ">>"C:\temp\certs\logs\reportecerts-$fecha.log"
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
