Sub IdentificarCodigosFaltantes()
    Dim wbData As Workbook, wbAero As Workbook
    Dim wsData As Worksheet, ws2D As Worksheet, ws3D As Worksheet
    Dim fileData As String, fileAero As String
    Dim lastRow As Long, lastRow2D As Long, lastRow3D As Long
    Dim dictAero As Object
    Dim i As Integer
    Dim codigo As String, fecha As Date
    Dim missingCodes As String, existingCodes As String
    Dim filePath As String
    Dim fileNumber As Integer
    Dim limite As Integer
    Dim countMissing As Integer, countExisting As Integer
    limite = 20 ' Máximo de códigos a mostrar en el MsgBox

    ' Seleccionar el archivo "Data Table.xlsx"
    fileData = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Selecciona el archivo Data Table")
    If fileData = "False" Then Exit Sub
    
    ' Seleccionar el archivo "Aero 2025 Test.xlsx"
    fileAero = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Selecciona el archivo Aero 2025 Test")
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

    ' Inicializar variables
    missingCodes = ""
    existingCodes = ""
    countMissing = 0
    countExisting = 0

    filePath = ThisWorkbook.Path & "\Missing_Codes.txt"
    fileNumber = FreeFile()
    Open filePath For Output As #fileNumber
    Print #fileNumber, "Códigos Faltantes (Missing Codes):"

    ' Filtrar por fechas en las próximas dos semanas y extraer códigos
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        If IsDate(wsData.Cells(i, 1).Value) Then
            fecha = CDate(wsData.Cells(i, 1).Value)
            If fecha >= Date And fecha <= Date + 14 Then
                codigo = Trim(Split(wsData.Cells(i, 5).Value, ":")(0)) ' Extraer código antes de ":"
                
                If Not dictAero.exists(codigo) Then
                    ' Código no encontrado -> Missing Codes
                    Print #fileNumber, codigo
                    countMissing = countMissing + 1
                    If countMissing <= limite Then
                        missingCodes = missingCodes & codigo & vbNewLine
                    ElseIf countMissing = limite + 1 Then
                        missingCodes = missingCodes & "..." & vbNewLine & "(Demasiados para mostrar en un solo mensaje)"
                    End If
                Else
                    ' Código encontrado -> Existing Codes
                    countExisting = countExisting + 1
                    If countExisting <= limite Then
                        existingCodes = existingCodes & codigo & vbNewLine
                    ElseIf countExisting = limite + 1 Then
                        existingCodes = existingCodes & "..." & vbNewLine & "(Demasiados para mostrar en un solo mensaje)"
                    End If
                End If
            End If
        End If
    Next i

    ' Agregar Existing Codes al archivo de texto
    Print #fileNumber, vbNewLine & "Códigos Encontrados en Aero (Existing Codes):"
    Print #fileNumber, existingCodes
    Close #fileNumber ' Cerrar el archivo

    ' Cerrar archivos sin guardar cambios
    wbData.Close False
    wbAero.Close False

    ' Mostrar los códigos en un MsgBox
    If countMissing = 0 Then
        missingCodes = "No hay códigos faltantes."
    End If
    If countExisting = 0 Then
        existingCodes = "No hay códigos que ya estuvieran."
    End If

    MsgBox "Códigos faltantes:" & vbNewLine & missingCodes & vbNewLine & vbNewLine & _
           "Códigos encontrados en Aero:" & vbNewLine & existingCodes & vbNewLine & vbNewLine & _
           "Se ha guardado un archivo llamado 'Missing_Codes.txt' con todos los códigos.", vbInformation, "Resultado"
End Sub
