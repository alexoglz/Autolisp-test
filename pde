Sub IdentificarCodigosFaltantes()
    Dim wbData As Workbook, wbAero As Workbook
    Dim wsData As Worksheet, ws2D As Worksheet, ws3D As Worksheet, wsFaltantes As Worksheet
    Dim fileData As String, fileAero As String
    Dim lastRow As Long, lastRow2D As Long, lastRow3D As Long, lastRowFaltantes As Long
    Dim dictAero As Object
    Dim i As Integer
    Dim codigo As String, fecha As Date
    
    ' Seleccionar el archivo "Data Table.xlsx"
    fileData = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Selecciona el archivo Data Table")
    If fileData = "False" Then Exit Sub
    
    ' Seleccionar el archivo "Aero.xlsx"
    fileAero = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Selecciona el archivo Aero")
    If fileAero = "False" Then Exit Sub
    
    ' Abrir los archivos seleccionados
    Set wbData = Workbooks.Open(fileData)
    Set wbAero = Workbooks.Open(fileAero)
    
    ' Seleccionar la primera hoja de Data Table
    Set wsData = wbData.Sheets(1)
    
    ' Seleccionar las hojas de Aero.xlsx
    Set ws2D = wbAero.Sheets("2D activities")
    Set ws3D = wbAero.Sheets("3D activities")
    
    ' Crear un diccionario para almacenar códigos de Aero
    Set dictAero = CreateObject("Scripting.Dictionary")
    
    ' Obtener los códigos de 2D activities (columna L)
    lastRow2D = ws2D.Cells(ws2D.Rows.Count, "L").End(xlUp).Row
    For i = 2 To lastRow2D
        If ws2D.Cells(i, "L").Value <> "" Then
            dictAero(ws2D.Cells(i, "L").Value) = 1
        End If
    Next i

    ' Obtener los códigos de 3D activities (columna H)
    lastRow3D = ws3D.Cells(ws3D.Rows.Count, "H").End(xlUp).Row
    For i = 2 To lastRow3D
        If ws3D.Cells(i, "H").Value <> "" Then
            dictAero(ws3D.Cells(i, "H").Value) = 1
        End If
    Next i

    ' Crear la hoja "Missing Codes" si no existe
    On Error Resume Next
    Set wsFaltantes = wbAero.Sheets("Missing Codes")
    If wsFaltantes Is Nothing Then
        Set wsFaltantes = wbAero.Sheets.Add
        wsFaltantes.Name = "Missing Codes"
    Else
        wsFaltantes.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Escribir encabezado
    wsFaltantes.Cells(1, 1).Value = "Códigos Faltantes"
    
    ' Obtener última fila de Data Table
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastRowFaltantes = 2

    ' Filtrar por fechas en las próximas dos semanas y extraer códigos
    For i = 2 To lastRow
        If IsDate(wsData.Cells(i, 1).Value) Then
            fecha = CDate(wsData.Cells(i, 1).Value)
            If fecha >= Date And fecha <= Date + 14 Then
                codigo = Trim(Split(wsData.Cells(i, 5).Value, ":")(0)) ' Extraer código antes de ":"
                If Not dictAero.exists(codigo) Then
                    wsFaltantes.Cells(lastRowFaltantes, 1).Value = codigo
                    lastRowFaltantes = lastRowFaltantes + 1
                End If
            End If
        End If
    Next i
    
    ' Cerrar archivo Data Table sin guardar
    wbData.Close False
    
    ' Guardar y cerrar Aero.xlsx
    wbAero.Save
    wbAero.Close True
    
    MsgBox "Proceso completado. Revisa la hoja 'Missing Codes' en Aero.xlsx.", vbInformation
End Sub
