Dim codigosFaltantes As String
Dim filePath As String
Dim limite As Integer
Dim fileNumber As Integer
limite = 20 ' Máximo de códigos a mostrar en el MsgBox

' Crear la lista de códigos faltantes
codigosFaltantes = ""
filePath = ThisWorkbook.Path & "\Missing_Codes.txt"

fileNumber = FreeFile()
Open filePath For Output As #fileNumber
Print #fileNumber, "Códigos Faltantes:"

For i = 2 To lastRowFaltantes - 1
    Print #fileNumber, wsFaltantes.Cells(i, 1).Value ' Escribir en el Bloc de Notas
    
    ' También agregar al mensaje emergente hasta 20 códigos
    If i <= limite Then
        codigosFaltantes = codigosFaltantes & wsFaltantes.Cells(i, 1).Value & vbNewLine
    ElseIf i = limite + 1 Then
        codigosFaltantes = codigosFaltantes & "..." & vbNewLine & "(Demasiados para mostrar en un solo mensaje)"
    End If
Next i

Close #fileNumber ' Cerrar el archivo

' Mostrar los códigos en un MsgBox
If codigosFaltantes = "" Then
    MsgBox "No hay códigos faltantes.", vbInformation, "Resultado"
Else
    MsgBox "Códigos faltantes encontrados:" & vbNewLine & codigosFaltantes & vbNewLine & vbNewLine & "Se ha guardado un archivo llamado 'Missing_Codes.txt' con todos los códigos.", vbInformation, "Resultado"
End If
