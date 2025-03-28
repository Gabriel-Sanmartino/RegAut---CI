<h1 style="text-align: center;">REGISTRO AUTOMATICO DE COLA DE IMPRESION</h1>

"RegAut - CI", es una solucion desarrollada a traves script, que tal como lo indica el titulo, permite generar un registro automatico de la cola de impresiones en un archivo 
CSV, esta desarrollada para usarse en todas las versiones de Windows.

¿En que entornos y como implementarla?

RegAut - CI, es una solucion pensada para utilizarse en ambitos laborales donde se imprime un alto volumen de acrhivos o documentos, esta permite tener un control o un registo
de las impresiones que realiza cada usuario o empleado.

Por ejemplo podria utilizarse para controlar o monitorear que los equipos se esten utilizando correctamente para imprimir solo documentos requiridos para las distintas labores dentro del entorno de trabajo, verificar que no se esten emitiendo documentos duplicados, entre otros usos.

Esto es posible ya que el registro automatico generado incluye detalles de fecha y hora, nombre del documento, usuario y cantidad de paginas impresas.
Ademas una ves que que el script se pone a funcionar, no es necesario volverlo a iniciar, esto es posible ya que esta configurado para iniciar cada ves que los equipos se encienden
lo cual evita la necesidad de destinar personal para la realizacion de esta tarea.

GLOSARIO: 
	
	RegAut - CI: Registro automatico de cola de impresion (carpeta).
	IniciarScript.vbs: Script que permite la ejecucion automatica.
	RegistrarColaImpresion.ps1: Script que registra la cola de impresion en archivo CSV.
	cola_impresion.csv: archivo donde se guardan los registros.

PASOS PARA PONER EN FUNCIONAMIENTO SCRIPTS:

1.Asignar permisos a PowerShell para la ejecucion de scripts:

	1.1.Ejecutar PowerShell como administrador
	
	1.2.Pegar el siguiente codigo en la terminal de PowerShell: 
		
		Set-ExecutionPolicy RemoteSigned -Scope LocalMachine

	1.3.En la terminal aparecera una pregunta, responder con la opcion "[O]" (letra), (en caso de que apareza otra opcion elegir "Si" o "Si a todo).
	
	1.4.Verificar cambio de configuracion, pegar el siguiente codigo en terminal PS:

		Get-ExecutionPolicy -List
	
	Debera imprimir por pantalla un mensaje similar al siguiente:
 
		```
		Scope            ExecutionPolicy
		---------------  ---------------
		MachinePolicy    Undefined
		UserPolicy       Undefined
		Process          Undefined
		CurrentUser      Undefined
		LocalMachine     RemoteSigned
		```



2.Reemplazar rutas en cada script:

	2.1.Reemplazar en script: "RegistrarColaImpresion.ps1":
		
		Abrir carpeta "RegAut - CI", click derecho al archivo "RegistrarColaImpresion.ps1" y luego seleccionar editar, en la segunda linea reemplazar
		con la ruta donde se guardo la carpeta "RegAut - CI", tal como se indica en el codigo.
	
		# Nombre del archivo de salida
		$outputFile = "C:\Reemplazar con la ruta donde se guardo la carpeta "RegAut - CI"\RegAut - CI\cola_impresion.csv"

	2.2.Reemplazar en script: "IniciarScript.vbs":

		En carpeta "RegAut - CI" click derecho, editar al archivo "IniciarScript.vbs", reemplazar ruta donde se indica en el codigo.

		Set WshShell = CreateObject("WScript.Shell")
		WshShell.Run "powershell -NoProfile -ExecutionPolicy Bypass -File ""C:\Reemplazar con la ruta donde se guardo la carpeta "RegAut - CI"\RegAut - CI			\RegistrarColaImpresion.ps1""", 0, False

3.Configurar ejecucion automatica y en segundo plano al encender el equipo:

	3.1. Presionar "Windows + R", escribir en el cuadro de dialogo: "shell:startup" (sin comillas) y presionar enter, esto abrira la carpeta inicio, pegar en ese
	direcctorio el archivo "IniciarScript.vbs"

	3.2. Reiniciar el ordenador, luego del reinicio el script "RegistrarColaImpresion.ps1" se automaticamente y en segundo plano, para verificar rapidamente
	ir a la carpeta "RegAut - CI", donde esta guardo el script, en dicha carpeta debera aparecer el nuevo archivo CSV "cola_impresion.csv", donde se registraran 
	todos los trabajos que pasen por cola de impresion.

NOTA: Si se desea detener el script, abrir administrador de tareas (Ctrl+Alt+Supr), pestaña detalles, proceso: "powershell.exe", click derecho, finalizar tarea.

Tecnologias utilizadas:

PowerShell: Lenguaje de scripts, diseñado para la automatizacion de tareas administrativas.
Visual Basic 6 (VB6): Lenguaje de programacion orientado a eventos.


