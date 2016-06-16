Attribute VB_Name = "ItemUpdater_module"
Sub ItemUpdater()
 On Error GoTo ErrorHandler
   
 Dim obj As Object
 Dim Sel As Outlook.Selection
 Dim strCategories As String
 Dim strSubject As String
 Dim strCurrentProject As String
 Dim intPosStart As Integer
 Dim intPosSlut As Integer
 
 Dim Categories() As String
 
 Set Sel = Application.ActiveExplorer.Selection
  If Sel.Count Then
   Set obj = Sel(1)
   If TypeOf obj Is Outlook.JournalItem Then
    MsgBox ("JournalItem")
   ElseIf TypeOf obj Is Outlook.MailItem Then
    strSubject = obj.Subject
                        

    '- Load form -------------------------------------------------------------------------------
    ItemDoneForm.Show

    '------------------------------------------------------------------------------------------
    '--- Updated settings ---------------------------------------------------------------------
    '------------------------------------------------------------------------------------------
    'Area Update
    'Get selected in listbox
    For i = 0 To ItemDoneForm.ListBox_Areas.ListCount - 1
     If ItemDoneForm.ListBox_Areas.Selected(i) Then
      
      For j = LBound(NewAreas, 1) To UBound(NewAreas, 1)
        If NewAreas(j, 0) = ItemDoneForm.ListBox_Areas.List(i) Then
         If strCategories <> "" Then strCategories = strCategories & ","
         strCategories = strCategories & NewAreas(j, 2)
        End If
      Next j
      
     End If
    Next i
    
    'Manufacturers Update
    'Get selected in listbox
    For i = 0 To ItemDoneForm.ListBox_Manufacturers.ListCount - 1
     If ItemDoneForm.ListBox_Manufacturers.Selected(i) Then
   
      For j = LBound(Manufacturers, 1) To UBound(Manufacturers, 1)
        If Manufacturers(j, 0) = ItemDoneForm.ListBox_Manufacturers.List(i) Then
         If strCategories <> "" Then strCategories = strCategories & ","
         strCategories = strCategories & Manufacturers(j, 2)
        End If
      Next j
      
     End If
    Next i
    
    'Status Update
    'Get selected in listbox
    For i = 0 To ItemDoneForm.ListBox_Status.ListCount - 1
     If ItemDoneForm.ListBox_Status.Selected(i) Then
   
      For j = LBound(Status, 1) To UBound(Status, 1)
        If Status(j, 0) = ItemDoneForm.ListBox_Status.List(i) Then
         If strCategories <> "" Then strCategories = strCategories & ","
         strCategories = strCategories & Status(j, 2)
        End If
      Next j
      
     End If
    Next i
            
                
'--- Projects control ----------------------------------------------------
           
    'Remove current project tag from subject
    If InStr(strSubject, "[RAP") <> 0 Then
     intPosStart = InStr(strSubject, "[RAP")
     intPosSlut = intPosStart + 8
     strCurrentProject = Mid(strSubject, intPosStart - 1, intPosSlut - intPosStart + 1)
     strSubject = Replace(strSubject, strCurrentProject, "")
    End If
                
    If InStr(strSubject, "[None]") <> 0 Then
     strSubject = Replace(strSubject, "[None]", "")
    End If

            
    If ItemDoneForm.ListBox_Projects.Value <> "" Then
     If strCategories <> "" Then strCategories = strCategories & ","
     strCategories = strCategories & NewProjects(ItemDoneForm.ListBox_Projects.ListIndex, 3)
     If InStr(strSubject, NewProjects(ItemDoneForm.ListBox_Projects.ListIndex, 0)) = 0 Then
      strSubject = strSubject & " [" & NewProjects(ItemDoneForm.ListBox_Projects.ListIndex, 0) & "]"
     End If
    End If
                 
                 
                 
                 
With obj
                    
 If ItemDoneForm_ButtonClicked = "Update" Then
  If ItemDoneForm.ComboBox_Due.ListIndex <> -1 Then
   .MarkAsTask olMarkThisWeek
   .TaskStartDate = Now
   .TaskDueDate = Now + ItemDoneForm.ComboBox_Due.Value
   '.ReminderSet = True
   '.ReminderTime = Now + 1
  End If
 End If
            
 If ItemDoneForm_ButtonClicked = "Done" Then
  .FlagStatus = olFlagComplete
  Categories = Split(strCategories, ",")
  For i = LBound(Categories) To UBound(Categories)
   Categories(i) = Trim(Categories(i))
   If InStr(Categories(i), "[{S") <> 0 Then
    strCategories = Replace(strCategories, Categories(i), "")
   End If
  Next
 End If
 
 
 .Categories = strCategories
 .Subject = strSubject
 
 .Save
End With
            
            
            
            ItemDoneForm_ButtonClicked = ""
            Unload ItemDoneForm
            
       
            
        Else
            MsgBox ("Warning: Unknown type")
        End If

    End If






    Exit Sub
     
ErrorHandler:
    MsgBox "HWHAP: " & vbNewLine & Err.Description

End Sub





