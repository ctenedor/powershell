' Ruta de la carpeta a la que deseas eliminar archivos
strRutaCarpeta = "C:\tmp"

' Crear un objeto FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Función para eliminar archivos de forma recursiva
Sub EliminarArchivosRecursivos(strRuta)
    ' Obtener la carpeta
    Set objCarpeta = objFSO.GetFolder(strRuta)

    ' Recorrer todos los archivos en la carpeta
    For Each objArchivo In objCarpeta.Files
        objArchivo.Delete True ' True para forzar la eliminación sin confirmación
    Next

    ' Recorrer todas las subcarpetas y llamar a la función de forma recursiva
    For Each objSubCarpeta In objCarpeta.SubFolders
        EliminarArchivosRecursivos objSubCarpeta.Path
    Next
End Sub

' Llamar a la función para eliminar archivos de forma recursiva
EliminarArchivosRecursivos strRutaCarpeta

' Liberar el objeto FileSystemObject
Set objFSO = Nothing
