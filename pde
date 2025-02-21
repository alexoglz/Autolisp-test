Sub IdentifyMissingCodes()
    Dim wbData As Workbook, wbAero As Workbook
    Dim wsData As Worksheet, ws2D As Worksheet, ws3D As Worksheet
    Dim fileData As String, fileAero As String
    Dim lastRow As Long, lastRow2D As Long, lastRow3D As Long
    Dim dictAero As Object
    Dim i As Integer
    Dim code As String, taskDate As Date
    Dim missingCodes As String, existingCodes As String
    Dim filePath As String
    Dim fileNumber As Integer
    Dim limit As Integer
    Dim countMissing As Integer, countExisting As Integer
    limit = 20 ' Maximum number of codes to display in the message box

    ' Select the "Data Table.xlsx" file
    fileData = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Select the Data Table file")
    If fileData = "False" Then Exit Sub
    
    ' Select the "Aero 2025 Test.xlsx" file
    fileAero = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Select the Aero 2025 Test file")
    If fileAero = "False" Then Exit Sub
    
    ' Open the selected files
    Set wbData = Workbooks.Open(fileData)
    Set wbAero = Workbooks.Open(fileAero)
    
    ' Select the first sheet in Data Table
    Set wsData = wbData.Sheets(1)
    
    ' Select the sheets from Aero 2025 Test
    Set ws2D = wbAero.Sheets("2D activities")
    Set ws3D = wbAero.Sheets("3D activities")
    
    ' Create a dictionary to store Aero 2025 Test codes
    Set dictAero = CreateObject("Scripting.Dictionary")
    
    ' Retrieve codes from 2D activities (Column L)
    lastRow2D = ws2D.Cells(ws2D.Rows.Count, "L").End(xlUp).Row
    For i = 2 To lastRow2D
        If ws2D.Cells(i, "L").Value <> "" Then
            dictAero(ws2D.Cells(i, "L").Value) = 1
        End If
    Next i

    ' Retrieve codes from 3D activities (Column H)
    lastRow3D = ws3D.Cells(ws3D.Rows.Count, "H").End(xlUp).Row
    For i = 2 To lastRow3D
        If ws3D.Cells(i, "H").Value <> "" Then
            dictAero(ws3D.Cells(i, "H").Value) = 1
        End If
    Next i

    ' Initialize variables
    missingCodes = ""
    existingCodes = ""
    countMissing = 0
    countExisting = 0

    filePath = ThisWorkbook.Path & "\Missing_Codes.txt"
    fileNumber = FreeFile()
    Open filePath For Output As #fileNumber
    Print #fileNumber, "Missing Codes:"

    ' Filter by dates within the next two weeks and extract codes
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        If IsDate(wsData.Cells(i, 1).Value) Then
            taskDate = CDate(wsData.Cells(i, 1).Value)
            If taskDate >= Date And taskDate <= Date + 14 Then
                code = Trim(Split(wsData.Cells(i, 5).Value, ":")(0)) ' Extract code before ":"
                
                If Not dictAero.exists(code) Then
                    ' Code not found -> Missing Codes
                    Print #fileNumber, code
                    countMissing = countMissing + 1
                    If countMissing <= limit Then
                        missingCodes = missingCodes & code & vbNewLine
                    ElseIf countMissing = limit + 1 Then
                        missingCodes = missingCodes & "..." & vbNewLine & "(Too many to display in one message)"
                    End If
                Else
                    ' Code found -> Existing Codes
                    countExisting = countExisting + 1
                    If countExisting <= limit Then
                        existingCodes = existingCodes & code & vbNewLine
                    ElseIf countExisting = limit + 1 Then
                        existingCodes = existingCodes & "..." & vbNewLine & "(Too many to display in one message)"
                    End If
                End If
            End If
        End If
    Next i

    ' Add Existing Codes to the text file
    Print #fileNumber, vbNewLine & "Existing Codes in Aero:"
    Print #fileNumber, existingCodes
    Close #fileNumber ' Close the text file

    ' Close files without saving changes
    wbData.Close False
    wbAero.Close False

    ' Display the results in a message box
    If countMissing = 0 Then
        missingCodes = "No missing codes."
    End If
    If countExisting = 0 Then
        existingCodes = "No codes were already present."
    End If

    MsgBox "Missing Codes:" & vbNewLine & missingCodes & vbNewLine & vbNewLine & _
           "Existing Codes in Aero:" & vbNewLine & existingCodes & vbNewLine & vbNewLine & _
           "A file named 'Missing_Codes.txt' has been saved with all the codes.", vbInformation, "Results"
End Sub
