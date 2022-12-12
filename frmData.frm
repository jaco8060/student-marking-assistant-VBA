VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmData 
   Caption         =   "Grading Assistant"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7305
   OleObjectBlob   =   "frmData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim filepath As String

Private Sub cmdBrowse_Click()
    
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.InitialFileName = ThisWorkbook.Path
    fd.Title = "Select a file."
    fd.AllowMultiSelect = False
    
    If fd.Show = -1 Then
        filepath = fd.SelectedItems(1)
    End If
    txtFilePath.Value = filepath
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdClear_Click()
    student_data.Range("A2", Range("A2").End(xlDown)).Value = ""
    student_data.Range("B2", Range("B2").End(xlDown)).Value = ""
    student_data.Range("C2", Range("C2").End(xlDown)).Value = ""
    
    student_data.Range("F2", Range("F2").End(xlDown)).Value = ""
    student_data.Range("G2", Range("G2").End(xlDown)).Value = ""
    
End Sub

Private Sub cmdOK_Click()
    Dim final_avg_list() As Variant
    Dim course_option As String
    
    Dim i As Integer
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    


        
    'create course_option string
    If CP102 Then
        course_option = "CP102"
    ElseIf CP212 Then
        course_option = "CP212"
    ElseIf CP104 Then
        course_option = "CP104"
    ElseIf AS101 Then
        course_option = "AS101"
    ElseIf PC120 Then
        course_option = "PC120"
    ElseIf PC131 Then
        course_option = "PC131"
    ElseIf PC141 Then
        course_option = "PC141"
    ElseIf CP411 Then
        course_option = "CP411"
    End If
    
    If enrollmentCMD Then
    
        SQL = "SELECT * FROM grades INNER JOIN students ON grades.studentID = students.studentID"
        
        i = 0
        
        'set active worksheet
        student_data.Activate
        
        If filepath <> "" Then
            ' Open the Connection string
            With cn
                .ConnectionString = "Data Source=" & filepath
                .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Open
            End With
            
            rs.Open SQL, cn
            
            With rs
            ' do while not end of file
                Do While Not .EOF
                    If .Fields("course") = course_option Then
                        student_data.Range("A2").Offset(i, 0) = .Fields("FirstName")
                        student_data.Range("B2").Offset(i, 0) = .Fields("LastName")
                        student_data.Range("C2").Offset(i, 0) = .Fields("grades.studentID")
                        i = i + 1
                    End If
                    ' Move the record pointer ahead one to get the next record
                    .MoveNext
                Loop
            End With

            
            'close connection and disassociate cn from connection object
            cn.Close
            
        Else
            txtFilePath.Value = "Please select a file."
            MsgBox ("Please select a file")
        End If

    
    ElseIf avgCMD Then
        SQL = "SELECT * FROM grades"
        
        'Activate sheet
        student_data.Activate
        
        i = 0
        If filepath <> "" Then
            ' Open the Connection string
            With cn
                .ConnectionString = "Data Source=" & filepath
                .Provider = "Microsoft.ACE.OLEDB.12.0"
                .Open
            End With
            
            rs.Open SQL, cn
            
            With rs
            ' do while not end of file
                Do While Not .EOF
                    If .Fields("course") = course_option Then
                        'add to lists with graded weights for avg calculations
                        ReDim Preserve final_avg_list(i)
                        
                        ' save data to worksheet
                        student_data.Range("F2").Offset(i, 0) = .Fields("studentID")
                        student_data.Range("G2").Offset(i, 0) = calc_avg(.Fields("A1"), .Fields("A2") _
                        , .Fields("A3"), .Fields("A4"), .Fields("MidTerm"), .Fields("Exam"))
                        
                        i = i + 1
                    End If
                    ' Move the record pointer ahead one to get the next record
                    .MoveNext
                Loop
            End With
            'close connection and disassociate cn from connection object
            
            cn.Close
            
            
            'refresh pivot table for pivot table/chart to update to new data values
            pivot_table.PivotTables("PivotTable1").PivotCache.refresh
            
            'output results to word document
            Call word_output
        Else
            txtFilePath.Value = "Please select a file."
            MsgBox ("Please select a file")
        End If

    End If

    
    

    
    

End Sub

'use to calculate final avg for each student
Function calc_avg(A1 As Integer, A2 As Integer, A3 As Integer, A4 As Integer, Midterm As Integer, Exam As Integer)
    calc_avg = ((A1 * 0.05) + (A2 * 0.05) + (A3 * 0.05) + (A4 * 0.05) + (Midterm * 0.3) + (Exam * 0.5))
End Function

'use to output to a word document

Sub word_output()
    
    Dim wdApp As Word.Application
    Dim wdDoc As Document
    Dim bookmark As Word.Range
    Dim bookmark2 As Word.Range
    'show word application
    Set wdApp = New Word.Application
    wdApp.Visible = True
Set wdDoc = wdApp.Documents.Open(ThisWorkbook.Path & "\Grade Assistant Word Document.docx")
    

    Set bookmark = wdDoc.Bookmarks("pivotTable").Range
    Set bookmark2 = wdDoc.Bookmarks("histogram").Range
    'delete previous results if there is any
    On Error Resume Next
    With wdDoc.InlineShapes(1)
        .Select
        .Delete
    End With
    On Error GoTo 0
    
    pivot_table.Range("Table1").Copy
    
    With bookmark
        .Select
        .PasteSpecial link:=False, DataType:=wdPasteBitmap, Placement:=wdInLine, DisplayAsIcon:=False
    End With
    
    
    pivot_table.ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Copy
    With bookmark2
        .Select
        .PasteSpecial link:=False, DataType:=wdPasteBitmap, Placement:=wdInLine, DisplayAsIcon:=False
    End With
    
    
End Sub



