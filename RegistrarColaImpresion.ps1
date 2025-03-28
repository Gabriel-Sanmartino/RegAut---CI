# Nombre del archivo de salida
$outputFile = "C:\Reemplazar con la ruta donde se guardo la carpeta RegAut - CI\RegAut - CI\cola_impresion.csv"

# Crear el archivo CSV con encabezados si no existe
if (!(Test-Path $outputFile)) {
    "Fecha y Hora,Nombre del Documento,Estado,Propietario,P�ginas,Tama�o,Enviado" | Out-File -FilePath $outputFile -Encoding utf8
}

# Historial para rastrear trabajos registrados
$historial = @{}

# Funci�n para registrar trabajos en la cola de impresi�n
function Registrar-ColaImpresion {
    # Consultar trabajos en la cola de impresi�n
    $trabajos = Get-WmiObject -Query "SELECT Document, Status, Owner, TotalPages, Size, TimeSubmitted FROM Win32_PrintJob"

    if ($trabajos) {
        foreach ($trabajo in $trabajos) {
            # Crear un identificador �nico basado en nombre del documento y hora enviada
            $identificadorUnico = "$($trabajo.Document)-$($trabajo.TimeSubmitted)"
            Write-Output "Procesando trabajo: $identificadorUnico"

            # Validaci�n: Mostrar valor de TotalPages antes de procesarlo
            Write-Output "TotalPages reportado por el sistema: $($trabajo.TotalPages)"

            # Verificar si el trabajo ya existe en el historial
            if ($historial.ContainsKey($identificadorUnico)) {
                # Actualizar el n�mero de p�ginas si el nuevo valor es mayor y v�lido
                if ($trabajo.TotalPages -gt $historial[$identificadorUnico].TotalPages -and $trabajo.TotalPages -ne 0) {
                    Write-Output "Actualizando p�ginas para: $identificadorUnico. Nuevo TotalPages: $($trabajo.TotalPages)"
                    $historial[$identificadorUnico].TotalPages = $trabajo.TotalPages
                } elseif ($trabajo.TotalPages -eq 0) {
                    Write-Output "Detectado TotalPages como 0. Manteniendo el valor anterior en el historial."
                } else {
                    Write-Output "El valor de TotalPages no es mayor ni v�lido. No se actualiza."
                }
            } else {
                # Agregar el nuevo trabajo al historial con el n�mero inicial de p�ginas
                $totalPagesFinal = if ($trabajo.TotalPages -gt 0) { $trabajo.TotalPages } else { 0 } # Si es 0, lo dejamos como tal
                Write-Output "A�adiendo nuevo trabajo al historial: $identificadorUnico con TotalPages: $totalPagesFinal"
                $historial[$identificadorUnico] = @{
                    TotalPages = $totalPagesFinal
                    LineaBase = "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")),$($trabajo.Document),$($trabajo.Status),$($trabajo.Owner),$totalPagesFinal,$($trabajo.Size),$($trabajo.TimeSubmitted)"
                }
            }

            # Escribir en el CSV sobrescribiendo si la fila ya existe
            $entrada = $historial[$identificadorUnico]
            $lineaActualizada = "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")),$($entrada.LineaBase.Split(',')[1]),$($entrada.LineaBase.Split(',')[2]),$($entrada.LineaBase.Split(',')[3]),$($entrada.TotalPages),$($entrada.LineaBase.Split(',')[5]),$($entrada.LineaBase.Split(',')[6])"
            Write-Output "Escribiendo l�nea actualizada en CSV: $lineaActualizada"

            # Leer y convertir el contenido actual del archivo CSV
            $contenidoCSV = Import-Csv -Path $outputFile | ForEach-Object { [PSCustomObject]$_ }
            $existeFila = $false

            # Buscar si ya existe una fila con el mismo Document y TimeSubmitted
            foreach ($fila in $contenidoCSV) {
                if ($fila."Nombre del Documento" -eq $trabajo.Document -and $fila.Enviado -eq $trabajo.TimeSubmitted) {
                    Write-Output "Fila existente encontrada. Sobrescribiendo..."
                    $fila."Fecha y Hora" = ((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
                    $fila."P�ginas" = $entrada.TotalPages
                    $existeFila = $true
                }
            }

            # Sobrescribir el CSV con las actualizaciones
            if ($existeFila) {
                Write-Output "Exportando contenido actualizado al archivo CSV."
                $contenidoCSV | Export-Csv -Path $outputFile -NoTypeInformation -Encoding utf8
            } else {
                # Si no existe la fila, agregarla al CSV
                Write-Output "A�adiendo nueva l�nea al archivo CSV."
                Add-Content -Path $outputFile -Value $lineaActualizada
            }
        }
    } else {
        Write-Output "No hay trabajos en la cola."
    }
}

# Monitorear la cola continuamente
while ($true) {
    Registrar-ColaImpresion
    Start-Sleep -Milliseconds 500 # Intervalo de monitoreo
}











