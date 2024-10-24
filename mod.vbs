' Crea un objeto FileSystemObject para manipular archivos y carpetas
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Obtiene la ruta de la carpeta actual (donde se encuentra el script)
Dim strCurrentFolder
strCurrentFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Establece la ruta de la carpeta "salida"
Dim strOutputFolder
strOutputFolder = strCurrentFolder & "\salida"

' Verifica si la carpeta "salida" existe
If Not objFSO.FolderExists(strOutputFolder) Then
    MsgBox "La carpeta 'salida' no existe. El script no puede continuar."
    WScript.Quit ' Detiene el script si no existe la carpeta
End If

' Obtiene una lista de todos los archivos .ui en la carpeta "salida"
Set objFolder = objFSO.GetFolder(strOutputFolder)
Set colFiles = objFolder.Files

Dim fileCount, processedFiles
fileCount = 0
processedFiles = 0

' Itera sobre cada archivo .ui en la carpeta "salida"
For Each objFile in colFiles
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "ui" Then
        fileCount = fileCount + 1
        On Error Resume Next ' Ignorar errores individuales de archivo
        ' Abre el archivo .ui para lectura
        Set objFileContent = objFSO.OpenTextFile(objFile.Path, 1) ' 1 = ForReading
        strFileContent = objFileContent.ReadAll
        objFileContent.Close

        ' Busca y reemplaza la cadena
        strNewContent = Replace(strFileContent, "<ui version=""4.0"">", "<ui version=""4.0"" language=""jambi"">", 1, 1)

        ' Crea la ruta del nuevo archivo .jui en la misma carpeta "salida"
        strNewFilePath = strOutputFolder & "\" & objFSO.GetBaseName(objFile.Name) & ".jui"

        ' Escribe el nuevo contenido en el archivo .jui (sobrescribe si existe)
        Set objFileContent = objFSO.OpenTextFile(strNewFilePath, 2, True) ' 2 = ForWriting
        objFileContent.WriteLine strNewContent
        objFileContent.Close
        
        ' Contar solo archivos procesados sin error
        If Err.Number = 0 Then
            processedFiles = processedFiles + 1
        End If
        On Error GoTo 0 ' Restablecer manejo de errores
    End If
Next

' Mensaje final según resultado
If fileCount > 0 Then
    If processedFiles = fileCount Then
        MsgBox "Se han procesado correctamente todos los archivos .ui."
    Else
        MsgBox "Hubo errores al procesar algunos archivos."
    End If
Else
    MsgBox "No se encontraron archivos .ui en la carpeta 'salida'."
End If
