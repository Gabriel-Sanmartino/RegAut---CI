Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell -NoProfile -ExecutionPolicy Bypass -File ""C:\Reemplazar con la ruta donde se guardo la carpeta RegAut - CI\RegAut - CI			\RegistrarColaImpresion.ps1""", 0, False
