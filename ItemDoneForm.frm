VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ItemDoneForm 
   Caption         =   "ItemDoneForm"
   ClientHeight    =   5556
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   9384
   OleObjectBlob   =   "ItemDoneForm_2016-06-16.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ItemDoneForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_Done_Click()
    ItemDoneForm_ButtonClicked = "Done"
    Me.Hide
End Sub

Private Sub Button_NextWeek_Click()
    ItemDoneForm_ButtonClicked = "Update"
    ItemDoneForm.ComboBox_Due.Value = 9 - Weekday(Now)
    Me.Hide
End Sub

Private Sub Button_Tester_Click()


End Sub

Private Sub Button_Today_Click()
    ItemDoneForm_ButtonClicked = "Update"
    ItemDoneForm.ComboBox_Due.Value = 0
    Me.Hide
End Sub

Private Sub Button_Tomorrow_Click()
    ItemDoneForm_ButtonClicked = "Update"
    ItemDoneForm.ComboBox_Due.Value = 1
    Me.Hide
End Sub

Private Sub Button_Update_Click()
    ItemDoneForm_ButtonClicked = "Update"
    Me.Hide
End Sub

Private Sub ListBox_Projects_Change()
    Debug.Print "THEREJ"
    'Area Update:
    For k = 0 To ListBox_Areas.ListCount - 1
     If InStr(Me.ListBox_Projects.List(Me.ListBox_Projects.ListIndex, 3), ListBox_Areas.List(k)) <> 0 Then
       ListBox_Areas.Selected(k) = True
     Else
      ListBox_Areas.Selected(k) = False
     End If
    Next k
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Prevent user from closing with the Close box in the title bar.
    If CloseMode <> 1 Then Cancel = 1
End Sub

Private Sub UserForm_Initialize()
 On Error GoTo ErrorHandler
 
 Dim tmpCount As Integer
 Dim intPosStart As Integer
 Dim intPosSlut As Integer
 Dim obj As Object
 Dim Sel As Outlook.Selection
 Dim tmpAreas As String
 Dim CategoryAreas() As String
 Dim CategoryManufacturers() As String
 Dim CategoryStatus() As String
 Dim strKeyWords As String
 Dim KWArray() As String
 
 'Load data
 'Load_Module.Load_Script

 
 
 ' ### ListBox AddItems -----------------------------------------------------------------------
  
 With ListBox_Projects
  .ColumnCount = 4
  .ColumnWidths = "40;;0;0"
  For tmpCount = LBound(NewProjects) To UBound(NewProjects)
   .AddItem NewProjects(tmpCount, 0)
   .Column(1, tmpCount) = NewProjects(tmpCount, 1)
   .Column(2, tmpCount) = NewProjects(tmpCount, 1)
   .Column(3, tmpCount) = NewProjects(tmpCount, 2)
  Next tmpCount
 End With
 
 With ListBox_Areas
  For tmpCount = LBound(NewAreas) To UBound(NewAreas)
   .AddItem NewAreas(tmpCount, 0)
  Next tmpCount
 End With
 
 With ListBox_Manufacturers
  For tmpCount = LBound(Manufacturers) To UBound(Manufacturers)
   .AddItem Manufacturers(tmpCount, 0)
  Next tmpCount
 End With
 
 With ListBox_Status
  For tmpCount = LBound(Status) To UBound(Status)
   .AddItem Status(tmpCount, 0)
  Next tmpCount
 End With
 
 
'------------------------------------------------------------------------------------------

 Set Sel = Application.ActiveExplorer.Selection
 If Sel.Count Then
  Set obj = Sel(1)
  If TypeOf obj Is Outlook.JournalItem Then
   MsgBox ("JournalItem")
  ElseIf TypeOf obj Is Outlook.MailItem Then
  
  
  

  
  
    
    ' ### Preset ListBox_Projects from Subject ##############################################
    ' Split Categories
    If InStr(obj.Subject, "[RAP") <> 0 Then
    intPosStart = InStr(obj.Subject, "[RAP")
    intPosSlut = intPosStart + 8
    For i = 0 To ItemDoneForm.ListBox_Projects.ListCount - 1
     If ItemDoneForm.ListBox_Projects.Column(0, i) = Mid(obj.Subject, intPosStart + 1, intPosSlut - intPosStart - 2) Then
    Debug.Print ItemDoneForm.ListBox_Projects.Column(0, i)
      
      ItemDoneForm.ListBox_Projects.Selected(i) = True
      Exit For
     End If
    Next
    End If
   '----------------------------------------------------------------------------------------------
    
    ' ### Preset ListBox_Status from Categories ##############################################
    ' Split Categories
    CategoryStatus = Split(obj.Categories, ";")
    
    For i = LBound(CategoryStatus) To UBound(CategoryStatus)
     CategoryStatus(i) = Trim(CategoryStatus(i))
     If InStr(CategoryStatus(i), "[{S") <> 0 Then
      intPosStart = InStr(CategoryStatus(i), "{S}")
      intPosSlut = InStr(CategoryStatus(i), "{/S}")
      CategoryStatus(i) = Mid(CategoryStatus(i), intPosStart + 3, intPosSlut - intPosStart - 3)
      For k = 0 To ListBox_Status.ListCount - 1
       If CategoryStatus(i) = ListBox_Status.List(k) Then ListBox_Status.Selected(k) = True
      Next
     End If
    Next
    ' ----------------------------------------------------------------------------------------
      
    ' ### Preset ListBox_Manufacturers from Categories ##############################################
    ' Split Categories
    CategoryManufacturers = Split(obj.Categories, ";")
    
    For i = LBound(CategoryManufacturers) To UBound(CategoryManufacturers)
     CategoryManufacturers(i) = Trim(CategoryManufacturers(i))
     If InStr(CategoryManufacturers(i), "[{M") <> 0 Then
      intPosStart = InStr(CategoryManufacturers(i), "{M}")
      intPosSlut = InStr(CategoryManufacturers(i), "{/M}")
      CategoryManufacturers(i) = Mid(CategoryManufacturers(i), intPosStart + 3, intPosSlut - intPosStart - 3)
      For k = 0 To ListBox_Manufacturers.ListCount - 1
       If CategoryManufacturers(i) = ListBox_Manufacturers.List(k) Then ListBox_Manufacturers.Selected(k) = True
      Next
     End If
    Next
    ' ----------------------------------------------------------------------------------------
    
    ' ### Preset ListBox_Manufacturers from KeyWords ##############################################
    ' Get Keywords
    For tmpCount = LBound(Manufacturers, 1) To UBound(Manufacturers, 1)
     
     strKeyWords = Manufacturers(tmpCount, 1)
     KWArray() = Split(strKeyWords, "|")
     For tmpCountA = LBound(KWArray) To UBound(KWArray)
      For k = 0 To ListBox_Manufacturers.ListCount - 1
       'Is Manufacturers keyword in subject
       If InStr(obj.Subject, KWArray(tmpCountA)) <> 0 Then
         If Manufacturers(tmpCount, 0) = ListBox_Manufacturers.List(k) Then ListBox_Manufacturers.Selected(k) = True
       End If
       'Is Manufacturers keyword in SenderEmailAddress
       If InStr(obj.SenderEmailAddress, KWArray(tmpCountA)) <> 0 Then
         If Manufacturers(tmpCount, 0) = ListBox_Manufacturers.List(k) Then ListBox_Manufacturers.Selected(k) = True
       End If
       'Is Manufacturers keyword in the Bodytext
       If InStr(obj.Body, KWArray(tmpCountA)) <> 0 Then
         If Manufacturers(tmpCount, 0) = ListBox_Manufacturers.List(k) Then ListBox_Manufacturers.Selected(k) = True
       End If
      Next k
     Next tmpCountA
    Next tmpCount
    ' ----------------------------------------------------------------------------------------
    
    
    ' ### Preset ListBox_Areas from Categories ##############################################
    ' Split Categories
    CategoryAreas = Split(obj.Categories, ";")
    
    For i = LBound(CategoryAreas) To UBound(CategoryAreas)
     CategoryAreas(i) = Trim(CategoryAreas(i))
     If InStr(CategoryAreas(i), "[{L") <> 0 Then
      intPosStart = InStr(CategoryAreas(i), "{L}")
      intPosSlut = InStr(CategoryAreas(i), "{/L}")
      CategoryAreas(i) = Mid(CategoryAreas(i), intPosStart + 3, intPosSlut - intPosStart - 3)
      For k = 0 To ListBox_Areas.ListCount - 1
       If CategoryAreas(i) = ListBox_Areas.List(k) Then ListBox_Areas.Selected(k) = True
      Next
     End If
    Next
    ' ----------------------------------------------------------------------------------------
   
    ' ### Preset ListBox_Areas from KeyWords ##############################################
    ' Get Keywords
    For tmpCount = LBound(NewAreas, 1) To UBound(NewAreas, 1)
     'Is area keyword in subject
     strKeyWords = NewAreas(tmpCount, 1)
     KWArray() = Split(strKeyWords, "|")
     For tmpCountA = LBound(KWArray) To UBound(KWArray)
      For k = 0 To ListBox_Areas.ListCount - 1
       If InStr(obj.Subject, KWArray(tmpCountA)) <> 0 Then
         If NewAreas(tmpCount, 0) = ListBox_Areas.List(k) Then ListBox_Areas.Selected(k) = True
       End If
       'Is Areas keyword in the Bodytext
       If InStr(obj.Body, KWArray(tmpCountA)) <> 0 Then
         If NewAreas(tmpCount, 0) = ListBox_Areas.List(k) Then ListBox_Areas.Selected(k) = True
       End If

      Next k
     Next tmpCountA
    Next tmpCount
    ' ----------------------------------------------------------------------------------------
   
   
  Else
   MsgBox ("Warning: Unknown type")
  End If
 End If




  
  With ComboBox_Due
    .AddItem "0"
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    .AddItem "10"
  End With
  
    Exit Sub
     
ErrorHandler:
    MsgBox "HWHAP: " & vbNewLine & Err.Description

End Sub


