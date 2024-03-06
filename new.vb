
Sub Test2()

    Dim strDocName As String
    Dim x As String
    Dim x1 As String
    Dim x2 As String
    Dim x3 As Long
    strDocName = ActiveDocument.Name
'    Set myRange = ActiveDocument.Content
'    Set aRange = ActiveDocument.Paragraphs(1).Style
    x = ActiveDocument.Paragraphs(3).Style
    x1 = ActiveDocument.Paragraphs(3).Range.Text
    x2 = ActiveDocument.Paragraphs(3).Parent
    x3 = ActiveDocument.Tables(1).Columns.Count
    
'    With aRange.Font
'     .Name = "Arial"
 '    .Size = 20
'    End With
    
    
'    ActiveDocument.Content.
    
    Debug.Print x
    Debug.Print x1, x3
'    ActiveDocument.Content.Bold = True
End Sub

Sub parse()
    Dim x As String
    Dim x1 As String
    Dim x2 As String
    Dim x3 As Long
    Dim x4 As String
'    Dim x5 As Object
    Dim MySize
'    Dim Cnxn As ADODB.Connection
'    Dim rstEmployees As ADODB.Recordset
'    Dim wr As System.IO.StreamWriter
   
    MySize = FileLen("C:\Users\ASUS\Documents\work_Romir\word\Aeceo a iaaacei_iiaay aa?ney_II Eeine_2024.03.docm")
'    Dim myArray(4) As String
    Dim celTable As Cell
   
    Set myTable = ActiveDocument.Tables(1)
    Set myRow = ActiveDocument.Tables(1).Rows(1)
    Debug.Print ActiveDocument.Name
'    Debug.Print ActiveDocument.Tables(1).Rows(1)
    x3 = ActiveDocument.Tables(1).Rows(2).Cells.Count
    x2 = myTable.Rows(3).Cells(2).Range.Text
    x4 = Trim(Left(x2, Len(x2) - 1))
   
    x2 = Replace(Replace(x2, Chr(10), ""), Chr(13), ". ")
    x2 = Left(x2, Len(x2) - 1)
   
    For Each celTable In ActiveDocument.Tables(1).Rows(5).Cells
'        Debug.Print ActiveDocument.Tables(1)
        Set my = celTable.Range
        Debug.Print my.Font.Size
'        Debug.Print celTable.Range.Parent
    Next celTable

    Set myNew = ActiveDocument.Tables(1).Rows(3).Cells(3).Range
    Debug.Print myNew.Paragraphs(1).Range.Text
   
'    Debug.Print "eie-ai eieiiie - ", x3
    Debug.Print "oaeno - ", x2
    Debug.Print "oaeno - ", x4, Len(x4), MySize
   
    Dim myArray(4) As String
   
    For i = 1 To myRow.Cells.Count
        x4 = myRow.Cells(i).Range.Text
'        Debug.Print I, Trim(Left(x4, Len(x4) - 1))
    Next i
   
    x4 = Replace(Replace(x4, Chr(10), ""), Chr(13), "")
   
    Open "C:\Users\ASUS\Documents\work_Romir\word\test.txt" For Output As #1
        Write #1, x2, "Hello world"
        Write #1, x5, " view"
    Close #1
   
'    Call test
   
End Sub
Sub test()
    Dim Cnxn As ADODB.Connection
    Dim rstEmployees As ADODB.Recordset
    Dim strCnxn As String
    Dim server_name As String
    Dim database_name As String
    Dim user_id As String
    Dim password As String
   
    Set Cnxn = New ADODB.Connection
   
    server_name = "localhost:3306" ' Enter your server name here - if running from a local computer use 127.0.0.1
    database_name = "lesson_4" ' Enter your database name here
    user_id = "root" ' enter your user ID here
    password = "7783Rafraikk@" ' Enter your password here
   
'    strCnxn = "Provider='sqloledb';Data Source='MySqlServer';" & _
'        "Initial Catalog='Northwind';Integrated Security='SSPI';"
       
'    Cnxn.Open strCnxn
   
'    Cnxn.Open "DRIVER={test}" _
'    & ";SERVER=" & server_name _
'    & ";DATABASE=" & database_name _
'    & ";user=" & user_id _
'    & ";password=" & password
   
    Cnxn.Open "Driver={MySQL ODBC 8.3 Driver};" & _
           "Server=127.0.0.1;" & _
           "Port=3306;" & _
           "Database=lesson_4;", "user=root;", "password=7783Rafraikk@;"
   
   
   
   
   
   
    Debug.Print "iiaay i?ioaao?a "
End Sub

Sub arrayTest()
    Dim arrT(1) As Integer
    Dim i As Integer, ii As Integer
   
    Dim A As Variant
'    A = Array()
    A = Array("wrer", "sd", "xcv")
'    A(0) = 3
'    A(1) = 4
    i = LBound(A) + UBound(A)
    ReDim Preserve A(i + 1)
    ii = LBound(A) + UBound(A)
    A(3) = "rwfsc"
'    A(1) = 20
    arrT(0) = 5
    arrT(1) = 7
    Debug.Print LBound(A)
    Debug.Print LBound(A)
    Debug.Print A(3), i, ii

End Sub
Sub arrayTest1()
    Dim A As Variant
    A = Array()
    i = LBound(A) + UBound(A)
    ReDim Preserve A(i + 1)
    A(i + 1) = "rwfsc"
    Debug.Print A(0)
    i = LBound(A) + UBound(A)
    ReDim Preserve A(i + 1)
    A(i + 1) = "cvbcbc"
    Debug.Print A(1)

End Sub


Public Sub mainStart()
'    Dim myTable As Table
    
    
    Debug.Print "Êîë-âî òàáëèö = ", ActiveDocument.Tables.Count
   
'    For Each myTable In ActiveDocument.Tables.
End Sub


Sub findMaxFontSize()
    Dim maxFontSize As Long, rowIndex As Long, flag As Long
    Dim myText As String
    Dim x1 As String
    Dim celTable As Cell
    Dim rowTable As Row
    Dim myParagraph As Paragraph
    Dim trimParagraph As Paragraph
   
    maxFontSize = 0
    rowIndex = 0
   
    For Each rowTable In ActiveDocument.Tables(1).Rows
        For Each celTable In rowTable.Cells
            For Each myParagraph In celTable.Range.Paragraphs
'                Set trimParagraph = myParagraph.Range.Words(1)
                If maxFontSize < myParagraph.Range.Words(1).Characters(1).Font.Size Then
                    maxFontSize = myParagraph.Range.Words(1).Characters(1).Font.Size
                    rowIndex = rowTable.Index
                    myText = myParagraph.Range.Text
                End If
            Next myParagraph
        Next celTable
'        Debug.Print rowTable.Index
    Next rowTable
    
    If ActiveDocument.Tables(1).Rows(rowIndex).Range.Paragraphs(1).Range.Words(1).Characters(1).Font.ColorIndex = -1 Then flag = 100
'    ActiveDocument.Tables(1).Rows(rowIndex).Range.Paragraphs(1).Range.Words(1).Characters(1).Font.Fill.Visible
'    flag = ActiveDocument.Tables(1).Rows(rowIndex).Range.Paragraphs(1).Range.Words(1).Characters(1).Font.Fill
    Debug.Print "Ìàêñèìàëüíûé øðèôò = ", maxFontSize, rowIndex, myText, flag

End Sub

Public Function CalculateSquareRoot(NumberArg As Double) As Double
 If NumberArg < 0 Then ' Evaluate argument.
  Exit Function ' Exit to calling procedure.
 Else
  CalculateSquareRoot = Sqr(NumberArg) ' Return square root.
 End If
End Function

Sub startParseTable()
    Dim maxFontSize As Long, rowIndex As Long, flag As Long
    Dim id As Long, sortingTitle As Long, sortingQuestion As Long
    Dim myText As String
    Dim x1 As String
    Dim celTable As Cell
    Dim rowTable As Row
    Dim myParagraph As String
    Dim trimParagraph As Paragraphs
    Dim textPart As Words
    Dim data As Variant, d As Variant
    Dim testMyString As String
    Dim indexCell As Integer
    
    data = Array()
    id = 1
    sortingTitle = 0
    sortingQuestion = 0
    
    For Each rowTable In ActiveDocument.Tables(1).Rows
        If rowTable.Index > 1 Then
            indexCell = 0
            For Each celTable In rowTable.Cells
                indexCell = indexCell + 1

'                For Each myParagraph In celTable.Range.Paragraphs
                    myParagraph = celTable.Range.Text
                    Set myParagraphs = celTable.Range.Paragraphs
                    If rowTable.Cells.Count = 1 Then
                    
                        ' Парсим название раздела
'                        Set myParagraphs = celTable.Range.Paragraphs
                        If myParagraphs(1).Range.Words(1).Characters(1).Font.Bold = -1 And myParagraphs(1).Range.Words(1).Characters(1).Font.Size = 12 Then
                            sortingTitle = sortingTitle + 1000
                            sortingQuestion = sortingTitle + 50
                            testMyString = chapterParseFunction(id, sortingTitle, myParagraph)
                            i = LBound(data) + UBound(data)
                            ReDim Preserve data(i + 1)
                            data(i + 1) = testMyString
                            Debug.Print testMyString, indexCell

                        End If
                    Else
                        If indexCell > 1 Then
                            ' Парсим вопрос
                            If indexCell = 2 Then
                                sortingQuestion = sortingQuestion + 50
                                testMyString = questionParseFunction(id, sortingQuestion, myParagraph, myParagraphs)
                                i = LBound(data) + UBound(data)
                                ReDim Preserve data(i + 1)
                                data(i + 1) = testMyString
                            End If
                            
                            ' Парсим ответ
                            If indexCell = 3 Then
'                                sortingQuestion = sortingQuestion + 50
                                testMyString = answerParseFunction(id, 8, myParagraph)
                                i = LBound(data) + UBound(data)
                                ReDim Preserve data(i + 1)
                                data(i + 1) = testMyString
                            End If
                        End If
                        
                    End If
                    
'                Next myParagraph
            id = id + 1
            Next celTable
        End If
        

'        sortingQuestion = sortingTitle + 50
    Next rowTable
    
    i = LBound(data) + UBound(data)
    
    Debug.Print "Ðàçìåð ìàññèâà = ", i + 1, data(1)
    
    For ii = 0 To i
        data(ii) = CStr(ii) + data(ii)
        Debug.Print ii, data(ii)
    Next ii
    
    Call writeToFile(data)


End Sub

Sub chapterParse(id As Long, sorting As Long, myParagraph As Paragraph)
    Dim myString As String, myText As String
    myText = myParagraph.Range.Text
    
    myText = Replace(Replace(myText, Chr(10), ""), Chr(13), ". ")
    myText = Left(myText, Len(myText) - 1)
    myText = Replace(myText, "..", ".")
    
'    myString = CStr(id) + "title" + ";" + CStr(sorting) + ";" + myText
'    myString = myText

    myString = CStr(id) + "title" + ";" + CStr(sorting) + ";" + myText
    Debug.Print myString

End Sub

Function chapterParseFunction(id As Long, sorting As Long, myParagraph As String) As String
    Dim myString As String, myText As String
'    myText = myParagraph.Range.Text
    myText = myParagraph
    
    myText = Replace(Replace(myText, Chr(10), ""), Chr(13), ". ")
    myText = Left(myText, Len(myText) - 1)
    myText = Replace(myText, "..", ".")
    
'    chapterParseFunction = CStr(id) + ";" + "title" + ";" + CStr(sorting) + ";" + myText + ";" + "null"
    chapterParseFunction = ";" + "title" + ";" + CStr(sorting) + ";" + myText + ";" + "null"

End Function

Function questionParseFunction(id As Long, sorting As Long, myParagraph As String, myParagraphs) As String
    Dim myString As String, myText As String
    Dim describe As String
    Dim indexParagraph As Integer
    
    describe = ""
    myText = ""
    indexParagraph = 1
'    Debug.Print myParagraphs.Count
    
    For Each mySentence In myParagraphs
        Debug.Print indexParagraph, mySentence.Range.Words(1).HighlightColorIndex
        If indexParagraph = myParagraphs.Count And mySentence.Range.Words(1).HighlightColorIndex = 3 Then
            describe = mySentence.Range.Text
        Else
            myText = myText + mySentence.Range.Text
        End If

        indexParagraph = indexParagraph + 1
    Next mySentence
    If describe = "" Then describe = "null"
'    myText = myParagraph.Range.Text
'    myText = myParagraph
    
    myText = Replace(Replace(myText, Chr(10), ""), Chr(13), ". ")
    myText = Left(myText, Len(myText) - 1)
    myText = Replace(myText, "..", ".")
    
    describe = Replace(Replace(describe, Chr(10), ""), Chr(13), ". ")
    describe = Left(describe, Len(describe) - 1)
    describe = Replace(describe, "..", ".")

    
    questionParseFunction = ";" + "question" + ";" + CStr(sorting) + ";" + myText + ";" + describe

End Function

Function answerParseFunction(id As Long, sorting As Long, myParagraph As String) As String
    Dim myString As String, myText As String
'    myText = myParagraph.Range.Text
    myText = myParagraph
    
    myText = Replace(Replace(myText, Chr(10), ""), Chr(13), ". ")
    myText = Left(myText, Len(myText) - 1)
    myText = Replace(myText, "..", ".")
    
    answerParseFunction = ";" + "answer" + ";" + CStr(sorting) + ";" + myText + ";" + "null"

End Function

Sub testFill()
    Debug.Print ActiveDocument.Tables(1).Rows(1).Range.Paragraphs(1).Range.Words(1).HighlightColorIndex
End Sub

Sub writeToFile(data As Variant)

    i = LBound(data) + UBound(data)
    Open "C:\Users\Abdyushev.R\Documents\VB_word\parse_table\test.txt" For Output As #1
        For ii = 0 To i
            Write #1, data(ii)
            
        Next ii
'        Write #1, x2, "Hello world"
'        Write #1, x5, " view"
    Close #1

End Sub
