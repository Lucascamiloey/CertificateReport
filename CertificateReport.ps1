
##
## Changelog:
##

# V 0.8.4
# 
#0.8
# Permite correr un subset definido de la lista que se encontro
# Permite usar AdFind para servidores que no permiten instalar RSAT 
# 
#0.8.1 
# AGREGA MANEJO DE ERRORES 
# 
#0.8.2 
# Agrega manejo de servers sin certificado (los cuenta y los muestra como vacios)
# 
#0.8.3 
# Cambia caracteres del log de errores para mejor manejo
# Agrega fila inicial en log de errores
# 
#0.8.4 
# Agrega filtro nuevo para rackspace
# Mejora el manejo de errores con errores mas explicitos
#
#0.8.5 
# Depreca Robocopy y evita la apertura de puertos de SMB
# Incrementa la velocidad de ejecucion un 70%



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
#Fija la fecha de inicio
$fechainicial=get-date 

## Setea el directorio de trabajo
$workingdir="c:\temp\certs\$fecha"

## Crea los directorios de trabajo ($workingdir y \logs)
New-Item -ItemType directory -Path $workingdir\logs -force
"ERROR;NAME">"$workingdir\logs\reportecerts-$fecha.log"


## Trae la lista de servers que tengan "server" en el campo operatingSystem 
## Con esto filtramos los clusters PERO incluimos los nodos.
## Query original 
## $Serverlist=(dsquery * -filter "(objectCategory=Computer)" -attr name operatingSystem -limit 10000)  | select @{label='ServerName';expression={(($_ -split ("   "))[0]).trim()}}, @{label='OperatingSystem';expression={(($_ -split ("   "))[1]  ).trim()}} | where {$_.operatingSystem -like "*server*"}
## Query modificada para rackspace US, solo servers Tier 1
$Serverlist=(dsquery * "OU=US,OU=Tier 1 Servers,OU=Servers,DC=discovery,DC=local" -filter "(objectCategory=Computer)" -attr name operatingSystem -limit 10000)  | select @{label='ServerName';expression={(($_ -split ("   "))[0]).trim()}}, @{label='OperatingSystem';expression={(($_ -split ("   "))[1]  ).trim()}} | where {$_.operatingSystem -like "*server*"}



#############################################################################################
## Tambien podemos usar AdFind (No necesita ser instalado, solo copiado dentro del %path%) ##
#############################################################################################
## Descomentar en caso de uso - Comentar la linea de DSQUERY

#  $serverlist=(AdFind.exe -f "(objectCategory=Computer)" Name operatingSystem -csv -nodn) | select @{label='ServerName';expression={(($_ -split (","))[0]).trim('"')}}, @{label='OperatingSystem';expression={(($_ -split (","))[1]  ).trim('"')}} | where {$_.operatingSystem -like "*server*"}

#############################################################################################
## Fin de uso de AdFind																	   ##	
#############################################################################################

## Exporta la lista a un csv para que se pueda ejecutar el proceso de recoleccion de certs
$serverlist | export-csv -notypeinformation $workingdir\servers.csv

############################################################################################
## Proceso para seleccionar un subset de la lista - descomentar en caso de querer usarlo  ##
############################################################################################
#
## Obtiene el subset de servers que se quieren procesar
# $acomparar=Get-Content $workingdir\subset.txt
### Crea el array de ayuda
# $newserverlist=@()
## Guarda el viejo $serverlist para futuras referencias
# $oldserverlist=$serverlist
## Compara todos los registros de $serverlist y guarda los que estan en el subset indicado en el array $newserverlist
# foreach ($item in $serverlist){
# 	if ($acomparar -contains $item.servername){
# 		$newserverlist+=$item
# 	}
# }
# 
#
## Hace el cambio de array para que el script siga funcionando adecuadamente con el nuevo subset
# $serverlist=$newserverlist
# 
############################################################################################
## Fin de uso de subset																	  ##
############################################################################################

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
		## WINRM
		#Invoca un comando remoto que trae los datos de los certificados, lo exporta a un CSV local, toma los datos crudos del CSV y los pega en un archivo local.
		Invoke-Command -ComputerName $name -ScriptBlock {Get-ChildItem -Path cert:\LocalMachine\My | Select-Object Subject, Issuer, NotBefore, NotAfter, Thumbprint, SerialNumber |   Export-Csv -path c:\tempcsv.csv -NoTypeInformation ; Get-Content C:\tempcsv.csv} | Set-Content $workingdir\$name.csv 2>"$workingdir\logs\temp.log" 
		
		# Descomentar la siguiente linea para debuggear el copiado
		write-output "Copiando $name.csv"
		
		## MANEJO de errores
		#Toma el temp.log 
		$temporal=get-content $workingdir\logs\temp.log
		#### SI ESTA TODO OK
		if ($temporal) 
		{			
			#si devuelve otra cosa, es un error a revisar
			$temporal>"$workingdir\logs\$name-Error-$fecha.log"
			"FATAL;$name">>"$workingdir\logs\reportecerts-$fecha.log"
			write-output "ERROR FATAL EN $name"
		}
	
	
	}else{
		#Graba error en $workingdir\logs\reportecerts-DIAMESAÑO-HORAMINUTO.log
		"OFFLINE;$name ">>"$workingdir\logs\reportecerts-$fecha.log"
		write-output "Timeout en $name"
	}
	
	# Marca la hora del final
	$fechafinal=get-date
}

## Recorre la carpeta buscando CSVs y guarda las rutas en $list
$list=Get-ChildItem $workingdir\* -Include *.csv -Exclude servers.csv

## Array de ayuda para recopilar certificados
$certhelper=@()
$certlist=@()
## Recorre el array $list y recopila la informacion en $certlist
foreach ($file in $list){
	$servername=$file.name.trim(".csv")
	$certhelper=Import-Csv $file | select @{label='ServerName';expression={$servername}}, *
	if ($certhelper)
	{
		$certlist+=$certhelper
	}else{
		$Tempcerthelp = New-Object PSCustomObject
		$Tempcerthelp | Add-Member -type NoteProperty -name ServerName -Value $servername
		$Tempcerthelp | Add-Member -type NoteProperty -name Subject -Value "None"
		$Tempcerthelp | Add-Member -type NoteProperty -name Issuer -Value "None"
		$Tempcerthelp | Add-Member -type NoteProperty -name NotBefore -Value "None"
		$Tempcerthelp | Add-Member -type NoteProperty -name NotAfter -Value "None"
		$Tempcerthelp | Add-Member -type NoteProperty -name Thumbprint -Value "0000"
		$Tempcerthelp | Add-Member -type NoteProperty -name SerialNumber -Value "0000"
		$certlist+=$Tempcerthelp
		Remove-Variable tempcerthelp		
	}
}

## Exporta el listado de certificados a CSV para que se pueda leer externamente
$certlist | export-csv -path $workingdir\FINAL-LIST.csv -notypeinformation

## Tiempo que tardo en hacer todos
$tiempototal=($fechafinal-$fechainicial).ToString().split(".")[0]

##########################################
## exportar datos a SQL: Pendiente 		##
## Evaluar el pase de fecha a juliano	##
## 		y castearlo como gregoriano 	##
## 		directamente en la base			##
##########################################

