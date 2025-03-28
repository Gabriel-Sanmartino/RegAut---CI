<h1 style="text-align: center;">✨🖨REGISTRO AUTOMATICO DE COLA DE IMPRESION🖨✨</h1>

**"RegAut - CI"**, es una solución desarrollada a través de scripts, que tal como lo indica el título, permite generar un registro ***automatico*** de la cola de impresiones en un archivo CSV, desarrollado para usarse en todas las versiones de Windows.

## 🔎¿En qué entornos y cómo implementarla? 

**RegAut - CI**, es una solución pensada 💡 para utilizarse en ámbitos laborales donde se imprime un alto volumen de archivos o documentos. Permite tener un control o un registro de las impresiones que realiza cada usuario o empleado.

**Por ejemplo**, podría utilizarse para controlar o monitorear que los equipos se estén utilizando correctamente para imprimir solo documentos requeridos para las distintas labores dentro del entorno de trabajo, verificar que no se estén emitiendo documentos duplicados, entre otros usos.

Esto es posible ya que ***el registro automático generado incluye detalles de fecha y hora, nombre del documento, usuario y cantidad de páginas impresas***. Además, una vez que el script se pone a funcionar, ***no es necesario volverlo a iniciar***, ya que está configurado para ejecutarse automáticamente al encender el equipo, ***evitando la necesidad de destinar personal para realizar esta tarea.***

**En resumen el uso del script implica una optimizacion en cuanto a tiempo y recursos** ⌛💰⚡

---

## 📑GLOSARIO: 
    
- **RegAut - CI**: Registro automático de cola de impresión (carpeta).
- **IniciarScript.vbs**: Script que permite la ejecución automática.
- **RegistrarColaImpresion.ps1**: Script que registra la cola de impresión en archivo CSV.
- **cola_impresion.csv**: Archivo donde se guardan los registros.

---

## 🔰PASOS PARA PONER EN FUNCIONAMIENTO LOS SCRIPTS:

### 1. Asignar permisos a PowerShell para la ejecución de scripts:

**1.1. Ejecutar PowerShell como administrador.**

**1.2. Pegar el siguiente código en la terminal de PowerShell:**

    ```powershell
    Set-ExecutionPolicy RemoteSigned -Scope LocalMachine
    ```

**1.3. En la terminal aparecerá una pregunta; responder con la opción `[O]` (letra). En caso de que aparezca otra opción, elegir "Sí" o "Sí a todo".**

**1.4. Verificar el cambio de configuración pegando el siguiente código en la terminal:**

    ```powershell
    Get-ExecutionPolicy -List
    ```

Debería imprimirse un mensaje similar al siguiente:

    ```
    Scope            ExecutionPolicy
    ---------------  ---------------
    MachinePolicy    Undefined
    UserPolicy       Undefined
    Process          Undefined
    CurrentUser      Undefined
    LocalMachine     RemoteSigned
    ```

---

### 2. Reemplazar rutas en cada script:

**2.1. Reemplazar en script: "RegistrarColaImpresion.ps1":**

Abrir la carpeta "RegAut - CI", hacer clic derecho en el archivo "RegistrarColaImpresion.ps1" y luego seleccionar editar. En la segunda línea, reemplazar con la ruta donde se guardó la carpeta "RegAut - CI", tal como se indica en el código.

```powershell
# Nombre del archivo de salida
$outputFile = "C:\Reemplazar con la ruta donde se guardó la carpeta 'RegAut - CI'\RegAut - CI\cola_impresion.csv"
```

**2.2. Reemplazar en script: "IniciarScript.vbs":**

En la carpeta "RegAut - CI", hacer clic derecho y editar el archivo "IniciarScript.vbs". Reemplazar la ruta donde se indica en el código.

```vbscript
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell -NoProfile -ExecutionPolicy Bypass -File ""C:\Reemplazar con la ruta donde se guardó la carpeta 'RegAut - CI'\RegAut - CI\RegistrarColaImpresion.ps1""", 0, False
```

---

### 3. Configurar ejecución automática y en segundo plano al encender el equipo:

3.1. Presionar `Windows + R`, escribir en el cuadro de diálogo: `shell:startup` (sin comillas) y presionar Enter. Esto abrirá la carpeta de inicio; pegar en ese directorio el archivo "IniciarScript.vbs".

3.2. Reiniciar el ordenador. Luego del reinicio, el script "RegistrarColaImpresion.ps1" se ejecutará automáticamente y en segundo plano. Para verificar rápidamente, ir a la carpeta "RegAut - CI", donde está guardado el script. En dicha carpeta deberá aparecer el nuevo archivo CSV "cola_impresion.csv", donde se registrarán todos los trabajos que pasen por la cola de impresión.

---

**NOTA:** Si se desea detener el script, abrir el administrador de tareas (`Ctrl + Alt + Supr`), pestaña "Detalles", buscar el proceso: "powershell.exe", hacer clic derecho y seleccionar "Finalizar tarea".

---

## 🚀Tecnologías utilizadas:

- **✅PowerShell**: Lenguaje de scripts diseñado para la automatización de tareas administrativas.
- **✅Visual Basic 6 (VB6)**: Lenguaje de programación orientado a eventos.

  <span style="color: red;">Texto en rojo</span>


🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵🔵
