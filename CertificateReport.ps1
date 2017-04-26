
##
## Changelog:
##
#
#0.8.8
#
# Suma reporte por mail.
# Envia una tabla con los certificados prontos a expirar (60 dias)
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
#Fija la fecha de inicio
$fechainicial=get-date 

## Setea el directorio de trabajo
$workingdir="c:\temp\certs\$fecha"

## Crea los directorios de trabajo ($workingdir y \logs)
New-Item -ItemType directory -Path $workingdir\logs -force
"ERROR;NAME">"$workingdir\logs\reportecerts-$fecha.log"

## Bloque de script que levanta los certificados
$ScriptBlock={Get-ChildItem -Path cert:\LocalMachine\My | Select-Object Subject, Issuer, NotBefore, NotAfter, Thumbprint, SerialNumber, @{label='DaysToExpiry';expression={[int]((get-date $_.NotAfter)-(get-date)).days}} | Export-Csv -path c:\tempcsv.csv -NoTypeInformation ; Get-Content C:\tempcsv.csv}

## Trae la lista de servers que tengan "server" en el campo operatingSystem 
## Con esto filtramos los clusters PERO incluimos los nodos.
## Query original 
## $Serverlist=(dsquery * -filter "(objectCategory=Computer)" -attr name operatingSystem -limit 10000)  | select @{label='ServerName';expression={(($_ -split ("   "))[0]).trim()}}, @{label='OperatingSystem';expression={(($_ -split ("   "))[1]  ).trim()}} | where {$_.operatingSystem -like "*server*"}
## Query modificada para rackspace US, solo servers Tier 1. 
$Serverlist=(dsquery * "OU=US,OU=Tier 1 Servers,OU=Servers,DC=discovery,DC=local" -filter "(objectCategory=Computer)" -attr name operatingSystem -limit 10000)  | select @{label='ServerName';expression={(($_ -split ("   "))[0]).trim()}}, @{label='OperatingSystem';expression={(($_ -split ("   "))[1]  ).trim()}} | where {$_.operatingSystem -like "*server*"}
## Agrega servers de PCE-A
$Serverlist+=(dsquery * "OU=PCE-A,OU=Tier 1 Servers,OU=Servers,DC=discovery,DC=local" -filter "(objectCategory=Computer)" -attr name operatingSystem -limit 10000)  | select @{label='ServerName';expression={(($_ -split ("   "))[0]).trim()}}, @{label='OperatingSystem';expression={(($_ -split ("   "))[1]  ).trim()}} | where {$_.operatingSystem -like "*server*"}



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
	if (test-wsman $name -ErrorAction silentlycontinue){
		## WINRM
		#Invoca un comando remoto que trae los datos de los certificados, lo exporta a un CSV local, toma los datos crudos del CSV y los pega en un archivo local.
		Invoke-Command -ComputerName $name -ScriptBlock $ScriptBlock 2>"$workingdir\logs\temp.log" | Set-Content $workingdir\$name.csv 
		## MANEJO de errores
		#Toma el temp.log 
		$temporal=get-content $workingdir\logs\temp.log
		#### SI NO ESTA TODO OK
		if ($temporal) 
		{			
			#si devuelve otra cosa, es un error a revisar
			$temporal>"$workingdir\logs\$name-Error-$fecha.log"
			"FATAL;$name">>"$workingdir\logs\reportecerts-$fecha.log"
			write-output "ERROR FATAL EN $name"
		}else{	
			# ESTA TODO PERFECTO
			# Descomentar la siguiente linea para debuggear el copiado
			write-output "Copiando $name.csv"
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
	#Chequea que el servidor tenga certificados
	$certhelper=Import-Csv $file | select @{label='ServerName';expression={$servername}}, *
	if ($certhelper)
	{
		$certlist+=$certhelper
	}else{
		# Si no tiene certificados, lo agrega igual para contabilizar
		$Tempcerthelp = New-Object PSCustomObject
		$Tempcerthelp | Add-Member -type NoteProperty -name ServerName -Value $servername
		$Tempcerthelp | Add-Member -type NoteProperty -name Subject -Value "None"
		$Tempcerthelp | Add-Member -type NoteProperty -name Issuer -Value "None"
		$Tempcerthelp | Add-Member -type NoteProperty -name NotBefore -Value "None"
		$Tempcerthelp | Add-Member -type NoteProperty -name NotAfter -Value "None"
		$Tempcerthelp | Add-Member -type NoteProperty -name Thumbprint -Value "0000"
		$Tempcerthelp | Add-Member -type NoteProperty -name SerialNumber -Value "0000"
		$Tempcerthelp | Add-Member -type NoteProperty -name DaysToExpiry -Value ""
		$certlist+=$Tempcerthelp
		Remove-Variable tempcerthelp		
	}
}

## Setea la cantidad de dias de aviso del mail
$dias=60

## Exporta el listado de certificados a CSV para que se pueda adjuntar
$certlist | export-csv -path $workingdir\FullList.csv -notypeinformation
$totalcerts=$certlist.count

## Exporta el listado de certificados  a menos de 60 dias de vencer a CSV para que se pueda adjuntar
$certlist | where {[int]$_.daystoexpiry -le $dias -and [int]$_.daystoexpiry -ge -10} | export-csv -path $workingdir\NearExpiry.csv -notypeinformation
$totalporexpirar=($certlist | where {[int]$_.daystoexpiry -le $dias -and [int]$_.daystoexpiry -ge -10}).count

## Trae la cantidad de errores ############ DEBUGGING
$errorlog=(import-csv $workingdir\logs\reportecerts-$fecha.log -Delimiter ";")
$offline=($errorlog | where {$_.error -eq "offline"}).count
$fatal=($errorlog | where {$_.error -eq "fatal"}).count

## Tiempo que tardo en hacer todos
$tiempototal=($fechafinal-$fechainicial).ToString().split(".")[0]

##########################################
## exportar datos a SQL: Pendiente 		##
## Evaluar el pase de fecha a juliano	##
## 		y castearlo como gregoriano 	##
## 		directamente en la base			##
##########################################

#############
## TESTING ##
#############

## Variables para email:

#############
### HEADER ##
#############

$Header = "<style>"
$Header = $Header + "BODY{background-color:white;}"
$Header = $Header + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$Header = $Header + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:gray}"
$Header = $Header + "TD{border-width: 1px;padding: 2px;border-style: solid;border-color: black;background-color:white}"
$Header = $Header + "</style>"

############
### BODY ###
############

$body=$certlist | where {[int]$_.daystoexpiry -le $dias -and [int]$_.daystoexpiry -ge -10} | select servername, subject, NotAfter, daystoexpiry | ConvertTo-Html -Head $Header | out-string

$toAddress="lucas.camilo@ar.ey.com"
$fromAddress="CertScript-TESTING@ey.com"
$subject="Certificate Report - $fecha"
$smtp="smtp.discovery.local"
$body+= "<br><br>
Se encontraron $totalcerts certificados, de los cuales $totalporexpirar expiran en los proximos $dias dias. Se encontraron $fatal errores y $offline equipos inalcanzables. <br>
Tardo $tiempototal hs en correr<br>
Adjuntos: Lista completa (FullList.csv), lista de proximos a expirar (NearExpiry.csv) y lista de errores (reportecerts-$fecha.log). <br> "
$Attachments="$workingdir\FullList.csv","$workingdir\NearExpiry.csv", "$workingdir\logs\reportecerts-$fecha.log"


# Send out the email message!
Send-MailMessage -to $toAddress  -from $fromAddress -subject $subject -smtpserver $smtp -body $body -Attachments $Attachments -BodyAsHtml 

