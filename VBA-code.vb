Slide100 - 1
Private Sub OptionButton1_Click()
End Sub
Private Sub OptionButton2_Click()
End Sub
Private Sub OptionButton3_Click()
End Sub
Private Sub CommandButton1_Click()
Dim crowd As String
If (OptionButton1 = True Or OptionButton2.Value = True Or OptionButton3.Value = True) Then
If OptionButton1.Value = True Then
crowd = "S"
ElseIf OptionButton2.Value = True Then
crowd = "DF"
ElseIf OptionButton3.Value = True Then
crowd = "DC"
End If
'MsgBox "crowd type is " & crowd
SaveToExcel (crowd)
nextest_slide
Else
MsgBox "Choose the type of crowd before proceeding"
End If
End Sub

Slide3 - 1
Private Sub CommandButton1_Click()
Dim sldTemp As Slide
Dim shpTemp As Shape
For Each sldTemp In ActivePresentation.Slides
For Each shpTemp In sldTemp.Shapes
If shpTemp.Type = msoOLEControlObject Then
If TypeName(shpTemp.OLEFormat.Object) = "OptionButton" Then
shpTemp.OLEFormat.Object.Value = False
End If
End If
Next
Next
ObserverData.Show
End Sub

Slide4 - 1
Private Sub CommandButton1_Click()
nextest_slide
End Sub

Slide7 - 1
Private Sub CommandButton1_Click()
Destroyer
ActivePresentation.SlideShowWindow.View.Exit
End Sub

Slide99 - 1
Private Sub CommandButton1_Click()
nextest_slide
End Sub

ObserverData - 1
Private Sub CommandButton1_Click()
If ObserverData.TextBox1.Value = "" Then
MsgBox "Enter Name Please"
ElseIf ObserverData.OptionButton1 = False And ObserverData.OptionButton2 = False Then
MsgBox "Select Age Please"
ElseIf ObserverData.OptionButton3 = False And ObserverData.OptionButton4 = False Then
MsgBox "Select gender Please"
Else
SetUserData
ActivePresentation.SlideShowWindow.View.Next
ObserverData.Hide
End If
SaveFormToExcel
End Sub
Private Sub Frame1_Click()
End Sub
Private Sub Frame2_Click()
End Sub
Private Sub Label1_Click()
End Sub
Private Sub OptionButton1_Click()
End Sub
Private Sub OptionButton2_Click()
End Sub

Module1 - 1
Sub nextest_slide()
ActivePresentation.SlideShowWindow.View.Next
End Sub

Module2 - 1
Dim column As String
Dim oXLApp As Object
Dim oWb As Object
Dim wb As Object
Public Sub SaveFormToExcel() 'ADDED
Set oWb = Nothing
Set oXLApp = Nothing
Dim col As Integer
Dim MyArray(26) As String
Set oXLApp = CreateObject("Excel.Application")
Set oWb = oXLApp.Workbooks.Open(ActivePresentation.Path & "\responses.xlsx")
If oWb.Worksheets(1).Range("A1") = "" Then
oWb.Worksheets(1).Range("A1") = "Name"
oWb.Worksheets(1).Range("A2") = "Age"
oWb.Worksheets(1).Range("A3") = "Gender"
End If
For intLoop = 0 To 25
MyArray(intLoop) = Chr$(64 + (intLoop + 1))
Next
col = 1
While oWb.Worksheets(1).Range(MyArray(col) & 1) <> ""
col = col + 1
Wend
column = MyArray(col)
oWb.Worksheets(1).Range(MyArray(col) & 1) = MyName
oWb.Worksheets(1).Range(MyArray(col) & 2) = MyAge
oWb.Worksheets(1).Range(MyArray(col) & 3) = MyGender
'skipped row count here, directly 1,2,3 here, also each time will try to name columns,
'video columns add manually in excel later iA-or code using slidenumber etc
oWb.Save
'oWb.Close
End Sub
Public Sub SaveToExcel(ByVal crowd As String) 'ADDED
Set wb = GetObject(ActivePresentation.Path & "\responses.xlsx")
'THIS is where I get problems. I do not want to OPEN the workbook since it is already open...
Dim row As Integer
'If oWb.Worksheets(1).Range(1 & (row)) = "" Then
' oWb.Worksheets(1).Range(1 & (row)) = slide_no
'End If
row = 4
While wb.Worksheets(1).Range(column & row) <> ""
row = row + 1

Module2 - 2
Wend
'column thru myarray set only once at start of slideshow, will change on restart of slideshow.
'should not be overwritten untill powerpoint is closed. the while loop will check for empty rows h
ere, and column
'in excel form. CANT PICK UP HALFWAY
wb.Worksheets(1).Range(column & row) = crowd
wb.Save
End Sub
Public Sub Destroyer()
wb.Close
Set oWb = Nothing
Set wb = Nothing
oXLApp.Quit
Set oXLApp = Nothing
End Sub
UserFormSample - 1
Public MyName As String
Public MyAge As String
Public MyGender As String
Dim strName As String
Dim strAge As String
Dim strGender As String
Sub SetUserData()
strName = ""
strAge = ""
strGender = ""
strName = ObserverData.TextBox1.Value
MyName = strName
If ObserverData.OptionButton1.Value = True Then
strAge = ObserverData.OptionButton1.Caption
MyAge = strAge
Else
strAge = ObserverData.OptionButton2.Caption
MyAge = strAge
End If
If ObserverData.OptionButton3.Value = True Then
strGender = ObserverData.OptionButton3.Caption
MyGender = strGender
Else
strGender = ObserverData.OptionButton4.Caption
MyGender = strGender
End If
End Sub