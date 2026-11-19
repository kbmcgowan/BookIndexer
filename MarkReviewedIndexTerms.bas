Attribute VB_Name = "MarkReviewedIndexTerms"
Option Explicit

Public Sub MarkReviewedIndexTerms()
    Dim csvPath As String
    Dim folderPath As String
    Dim csvDocument As Document
    Dim sourceDocument As Document
    Dim csvText As String
    Dim lines As Variant
    Dim fields As Collection
    Dim headerFound As Boolean
    Dim documentColumn As Long
    Dim termColumn As Long
    Dim keepColumn As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim documentName As String
    Dim term As String
    Dim keepValue As String
    Dim sourceNames As New Collection
    Dim rowDocuments As New Collection
    Dim rowTerms As New Collection
    Dim termsForDocument As Collection
    Dim sourceName As Variant
    Dim outputName As String
    Dim outputPath As String
    Dim separatorPosition As Long
    Dim markedTermCount As Long
    Dim answer As VbMsgBoxResult

    With Application.FileDialog(3)
        .Title = "Select the reviewed index-candidate CSV"
        .Filters.Clear
        .Filters.Add "CSV files", "*.csv"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        csvPath = .SelectedItems(1)
    End With

    separatorPosition = InStrRev(csvPath, "/")
    If InStrRev(csvPath, "\") > separatorPosition Then
        separatorPosition = InStrRev(csvPath, "\")
    End If
    If separatorPosition = 0 Then
        MsgBox "Could not determine the CSV's folder.", vbExclamation
        Exit Sub
    End If
    folderPath = Left$(csvPath, separatorPosition)

    Set csvDocument = Documents.Open( _
        FileName:=csvPath, _
        ConfirmConversions:=False, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        Format:=wdOpenFormatText, _
        Encoding:=65001)

    csvText = Replace$(csvDocument.Content.Text, vbCrLf, vbCr)
    csvText = Replace$(csvText, vbLf, vbCr)
    lines = Split(csvText, vbCr)

    For rowIndex = LBound(lines) To UBound(lines)
        If Len(Trim$(CStr(lines(rowIndex)))) > 0 Then
            Set fields = ParseCsvLine(CStr(lines(rowIndex)))

            If Not headerFound Then
                For columnIndex = 1 To fields.Count
                    Select Case LCase$(Trim$(Replace$( _
                        CStr(fields(columnIndex)), ChrW(&HFEFF), vbNullString)))
                        Case "document"
                            documentColumn = columnIndex
                        Case "candidate_term"
                            termColumn = columnIndex
                        Case "keep"
                            keepColumn = columnIndex
                    End Select
                Next columnIndex

                If documentColumn = 0 Or termColumn = 0 Or keepColumn = 0 Then
                    csvDocument.Close SaveChanges:=False
                    MsgBox "The CSV must have 'document', 'candidate_term', and " & _
                        "'Keep' columns. Review and save the CSV first.", vbExclamation
                    Exit Sub
                End If
                headerFound = True
            Else
                If fields.Count < documentColumn Or fields.Count < termColumn _
                    Or fields.Count < keepColumn Then
                    csvDocument.Close SaveChanges:=False
                    MsgBox "A CSV row has fewer columns than the header.", vbExclamation
                    Exit Sub
                End If

                documentName = Trim$(CStr(fields(documentColumn)))
                term = Trim$(CStr(fields(termColumn)))
                keepValue = LCase$(Trim$(CStr(fields(keepColumn))))

                Select Case keepValue
                    Case "", "false", "0", "no"
                        keepValue = "false"
                    Case "true", "1", "yes"
                        keepValue = "true"
                    Case Else
                        csvDocument.Close SaveChanges:=False
                        MsgBox "Invalid Keep value in CSV row " & CStr(rowIndex + 1) & _
                            ": " & keepValue, vbExclamation
                        Exit Sub
                End Select

                If keepValue = "true" Then
                    If Len(documentName) = 0 Or Len(term) = 0 Then
                        csvDocument.Close SaveChanges:=False
                        MsgBox "A selected CSV row needs both a document and a term.", _
                            vbExclamation
                        Exit Sub
                    End If
                    If InStr(documentName, "/") > 0 Or InStr(documentName, "\") > 0 _
                        Or LCase$(Right$(documentName, 5)) <> ".docx" Then
                        csvDocument.Close SaveChanges:=False
                        MsgBox "Invalid DOCX filename in CSV: " & documentName, _
                            vbExclamation
                        Exit Sub
                    End If

                    rowDocuments.Add documentName
                    rowTerms.Add term
                    If Not CollectionContainsText(sourceNames, documentName) Then
                        sourceNames.Add documentName
                    End If
                End If
            End If
        End If
    Next rowIndex

    csvDocument.Close SaveChanges:=False

    If Not headerFound Or rowDocuments.Count = 0 Then
        MsgBox "No candidate rows with Keep set to True were found in the CSV.", _
            vbExclamation
        Exit Sub
    End If

    For Each sourceName In sourceNames
        outputName = Left$(CStr(sourceName), Len(CStr(sourceName)) - 5) & "_indexed.docx"
        outputPath = folderPath & outputName

        If Len(Dir$(folderPath & CStr(sourceName))) = 0 Then
            MsgBox "Source document not found: " & folderPath & CStr(sourceName), _
                vbExclamation
            Exit Sub
        End If
        If Len(Dir$(outputPath)) > 0 Then
            MsgBox "Output already exists; no files were processed:" & vbCrLf & _
                outputPath & vbCrLf & _
                "Move or rename the existing indexed copy before running again.", _
                vbExclamation
            Exit Sub
        End If
    Next sourceName

    answer = MsgBox( _
        "This will mark every match for the reviewed CSV terms and save new " & _
        "'_indexed.docx' copies beside the originals. The original documents " & _
        "will not be changed." & vbCrLf & vbCrLf & _
        "Documents: " & sourceNames.Count & vbCrLf & _
        "Selected term rows: " & rowTerms.Count & vbCrLf & vbCrLf & _
        "Continue?", _
        vbQuestion + vbYesNo)
    If answer <> vbYes Then Exit Sub

    For Each sourceName In sourceNames
        Set sourceDocument = Documents.Open( _
            FileName:=folderPath & CStr(sourceName), _
            ConfirmConversions:=False, _
            ReadOnly:=True, _
            AddToRecentFiles:=False)
        sourceDocument.Activate
        Set termsForDocument = New Collection

        For rowIndex = 1 To rowDocuments.Count
            If StrComp(CStr(rowDocuments(rowIndex)), CStr(sourceName), vbTextCompare) = 0 Then
                term = CStr(rowTerms(rowIndex))
                If Not CollectionContainsText(termsForDocument, term) Then
                    termsForDocument.Add term
                    Application.MarkAllEntries Range:=sourceDocument.Content, Entry:=term
                    markedTermCount = markedTermCount + 1
                End If
            End If
        Next rowIndex

        outputName = Left$(CStr(sourceName), Len(CStr(sourceName)) - 5) & "_indexed.docx"
        outputPath = folderPath & outputName
        sourceDocument.SaveAs2 _
            FileName:=outputPath, _
            FileFormat:=wdFormatXMLDocument, _
            AddToRecentFiles:=False
        sourceDocument.Close SaveChanges:=False
    Next sourceName

    MsgBox "Finished. Created indexed copies for " & sourceNames.Count & _
        " documents and marked " & markedTermCount & _
        " unique term/document pairs." & vbCrLf & _
        "Review the copies before combining or generating the final index.", _
        vbInformation
End Sub

Private Function ParseCsvLine(ByVal line As String) As Collection
    Dim fields As New Collection
    Dim fieldValue As String
    Dim character As String
    Dim inQuotes As Boolean
    Dim position As Long

    For position = 1 To Len(line)
        character = Mid$(line, position, 1)

        If character = """" Then
            If inQuotes And Mid$(line, position + 1, 1) = """" Then
                fieldValue = fieldValue & """"
                position = position + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf character = "," And Not inQuotes Then
            fields.Add fieldValue
            fieldValue = vbNullString
        Else
            fieldValue = fieldValue & character
        End If
    Next position

    fields.Add fieldValue
    Set ParseCsvLine = fields
End Function

Private Function CollectionContainsText( _
    ByVal values As Collection, _
    ByVal sought As String) As Boolean
    Dim value As Variant

    For Each value In values
        If StrComp(CStr(value), sought, vbTextCompare) = 0 Then
            CollectionContainsText = True
            Exit Function
        End If
    Next value
End Function
