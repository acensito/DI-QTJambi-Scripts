' Crea un objeto FileSystemObject para manipular archivos y carpetas
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Obtiene la ruta de la carpeta actual (donde se encuentra el script)
Dim strCurrentFolder
strCurrentFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Crea la subcarpeta "salida" si no existe
Dim strOutputFolder
strOutputFolder = strCurrentFolder & "\salida"
If Not objFSO.FolderExists(strOutputFolder) Then
    objFSO.CreateFolder strOutputFolder
End If

' Obtiene una lista de todos los archivos .ui en la carpeta actual
Set objFolder = objFSO.GetFolder(strCurrentFolder)
Set colFiles = objFolder.Files

' Itera sobre cada archivo .ui
For Each objFile in colFiles
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "ui" Then
        ' Abre el archivo .ui para lectura
        Set objFileContent = objFSO.OpenTextFile(objFile.Path, 1) ' 1 = ForReading
        strFileContent = objFileContent.ReadAll
        objFileContent.Close

        ' Busca y reemplaza la cadena
        strNewContent = Replace(strFileContent, "<ui version=""4.0"">", "<ui version=""4.0"" language=""jambi"">", 1, 1)

        ' Crea la ruta del nuevo archivo .jui
        strNewFilePath = strOutputFolder & "\" & objFSO.GetBaseName(objFile.Name) & ".jui"

        ' Escribe el nuevo contenido en el archivo .jui (sobrescribe si existe)
        Set objFileContent = objFSO.OpenTextFile(strNewFilePath, 2, True) ' 2 = ForWriting
        objFileContent.WriteLine strNewContent
        objFileContent.Close
    End If
Next

MsgBox "Se han creado los archivos .jui en la carpeta 'salida'."