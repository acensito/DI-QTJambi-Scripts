Option Explicit

Dim objFSO, objFolder, objFile, strFolderPath, strCommand, objShell, intFileCount

' Crear un objeto FileSystemObject para trabajar con archivos
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
intFileCount = 0 ' Contador de archivos procesados

' Definir la ruta a la carpeta que contiene los archivos .jui (relativa a la ubicación del script)
strFolderPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\salida"

' Comprobar si la carpeta existe
If objFSO.FolderExists(strFolderPath) Then
    ' Obtener la carpeta
    Set objFolder = objFSO.GetFolder(strFolderPath)
    
    ' Recorrer todos los archivos en la carpeta
    For Each objFile In objFolder.Files
        ' Comprobar si la extensión es .jui
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "jui" Then
            ' Construir el comando para ejecutar juic.exe con el archivo .jui y la opción -d para la carpeta de salida
            strCommand = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\juic.exe -pf """ & objFile.Path & """ -d """ & strFolderPath & """"
            ' Ejecutar el comando
            objShell.Run strCommand, 0, True
            intFileCount = intFileCount + 1 ' Aumentar el contador de archivos procesados
        End If
    Next
    
    ' Mostrar mensaje final solo si se procesaron archivos
    If intFileCount > 0 Then
        WScript.Echo "Todos los archivos .jui fueron procesados correctamente. Total: " & intFileCount
    Else
        WScript.Echo "No se encontraron archivos .jui en la carpeta."
    End If
Else
    WScript.Echo "La carpeta 'salida' no existe: " & strFolderPath
End If

' Limpiar objetos
Set objFolder = Nothing
Set objFSO = Nothing
Set objShell = Nothing
