Sub IdentifyMissingCodes()
    ' Declare variables for the workbooks and sheets
    Dim wbData As Workbook, wbAero As Workbook
    Dim wsData As Worksheet, ws2D As Worksheet, ws3D As Worksheet, wsQuality As Worksheet
    
    ' Variables for file selection
    Dim fileData As String, fileAero As String
    
    ' Variables to track the last row in each sheet
    Dim lastRow As Long, lastRow2D As Long, lastRow3D As Long, lastRowQuality As Long
    
    ' Dictionary to store all codes found in Aero 2025 Test.xlsx
    Dim dictAero As Object
    
    ' Variables for looping and processing
    Dim i As Integer
    Dim code As String, taskDate As Date
    
    ' Variables to store missing and existing codes
    Dim missingCodes As String, existingCodes As String
    
    ' Variables for writing to a text file
    Dim filePath As String
    Dim fileNumber As Integer
    
    ' Limits the number of codes shown in the pop-up message to avoid flooding the user
    Dim limit As Integer
    Dim countMissing As Integer, countExisting As Integer
    limit = 20 ' Max number of codes to show in the MsgBox
    
    ' Prompt the user to select the Data Table file (this is the file that contains the task list)
    fileData = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Select the Data Table file")
    If fileData = "False" Then Exit Sub
    
    ' Now ask for the Aero 2025 Test file (this is the one where we check if the codes exist)
    fileAero = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Select the Aero 2025 Test file")
    If fileAero = "False" Then Exit Sub
    
    ' Open both selected files
    Set wbData = Workbooks.Open(fileData)
    Set wbAero = Workbooks.Open(fileAero)
    
    ' Grab the first sheet from Data Table (this is where we get the task descriptions)
    Set wsData = wbData.Sheets(1)
    
    ' Grab the 2D, 3D, and Quality Issues 3D sheets from Aero 2025 Test
    Set ws2D = wbAero.Sheets("2D activities")
    Set ws3D = wbAero.Sheets("3D activities")
    Set wsQuality = wbAero.Sheets("Quality Issues 3D")
    
    ' Create a dictionary to store all the codes from Aero 2025 Test
    Set dictAero = CreateObject("Scripting.Dictionary")
    
    ' First, collect all the codes from the 2D activities sheet (Column L)
    lastRow2D = ws2D.Cells(ws2D.Rows.Count, "L").End(xlUp).Row
    For i = 2 To lastRow2D
        If ws2D.Cells(i, "L").Value <> "" Then
            dictAero(ws2D.Cells(i, "L").Value) = 1
        End If
    Next i

    ' Now, grab all the codes from the 3D activities sheet (Column H)
    lastRow3D = ws3D.Cells(ws3D.Rows.Count, "H").End(xlUp).Row
    For i = 2 To lastRow3D
        If ws3D.Cells(i, "H").Value <> "" Then
            dictAero(ws3D.Cells(i, "H").Value) = 1
        End If
    Next i

    ' Finally, collect all the codes from the Quality Issues 3D sheet (Column H)
    lastRowQuality = wsQuality.Cells(wsQuality.Rows.Count, "H").End(xlUp).Row
    For i = 2 To lastRowQuality
        If wsQuality.Cells(i, "H").Value <> "" Then
            dictAero(wsQuality.Cells(i, "H").Value) = 1
        End If
    Next i

    ' Initialize variables for storing results
    missingCodes = ""
    existingCodes = ""
    countMissing = 0
    countExisting = 0

    ' Ask the user where they want to save the Missing_Codes.txt file
    filePath = Application.GetSaveAsFilename(InitialFileName:="Missing_Codes.txt", FileFilter:="Text Files (*.txt), *.txt", Title:="Choose where to save the file")
    
    ' If the user cancels, exit the macro
    If filePath = "False" Then Exit Sub
    
    ' Open the text file for writing
    fileNumber = FreeFile()
    Open filePath For Output As #fileNumber
    Print #fileNumber, "Missing Codes:"

    ' Now, let's go through the Data Table and check each task
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        ' Make sure the first column actually contains a valid date
        If IsDate(wsData.Cells(i, 1).Value) Then
            taskDate = CDate(wsData.Cells(i, 1).Value) ' Convert it to VBA date format
            
            ' Only process tasks that are within the next two weeks
            If taskDate >= Date And taskDate <= Date + 14 Then
                ' Extract the code from Column E (everything before the ":")
                code = Trim(Split(wsData.Cells(i, 5).Value, ":")(0)) 
                
                ' Check if this code exists in Aero 2025 Test
                If Not dictAero.exists(code) Then
                    ' If it's missing, add it to the text file and message
                    Print #fileNumber, code
                    countMissing = countMissing + 1
                    If countMissing <= limit Then
                        missingCodes = missingCodes & code & vbNewLine
                    ElseIf countMissing = limit + 1 Then
                        missingCodes = missingCodes & "..." & vbNewLine & "(Too many to display in one message)"
                    End If
                Else
                    ' If it's found, add it to the existing codes list
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

    ' Also print the existing codes in the text file
    Print #fileNumber, vbNewLine & "Existing Codes in Aero:"
    Print #fileNumber, existingCodes
    Close #fileNumber ' Done writing to the file, close it

    ' Save both files before finishing the macro
    wbData.Save
    wbAero.Save

    ' Show a pop-up with the results (but limit to 20 codes for readability)
    If countMissing = 0 Then
        missingCodes = "No missing codes."
    End If
    If countExisting = 0 Then
        existingCodes = "No codes were already present."
    End If

    MsgBox "Missing Codes:" & vbNewLine & missingCodes & vbNewLine & vbNewLine & _
           "Existing Codes in Aero:" & vbNewLine & existingCodes & vbNewLine & vbNewLine & _
           "The file has been saved at: " & filePath, vbInformation, "Results"
End Sub
