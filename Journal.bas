Attribute VB_Name = "Journal"
Public Sub CreateNewFromCurrentItem()
  Dim obj As Object
  Dim Sel As Outlook.Selection
  Dim J1 As Outlook.JournalItem
  Dim J2 As Outlook.JournalItem
  Dim Links1 As Outlook.Links
  Dim Links2 As Outlook.Links
  Dim T1 As Outlook.TaskItem
  Dim T2 As Outlook.TaskItem
  Dim i&
  Dim Link As Outlook.Link
  Dim F As Outlook.MAPIFolder

  Set Sel = Application.ActiveExplorer.Selection
  If Sel.Count Then
    Set obj = Sel(1)
    Set F = obj.Parent

    If TypeOf obj Is Outlook.JournalItem Then
      Set J1 = obj
      Set J2 = F.Items.Add(olJournalItem)

      With J1
        J2.Categories = .Categories
        J2.Companies = .Companies
        J2.ContactNames = .ContactNames
        J2.Start = Now
        J2.Subject = .Subject
        J2.Type = .Type
        J2.Body = .Body
      End With

      Set Links1 = J1.Links
      Set Links2 = J2.Links

      On Error Resume Next
      For i = 1 To Links1.Count
        Set Link = Links1(i)
        Links2.Add Link.Item
      Next

      J2.Save
      J2.Display

    ElseIf TypeOf obj Is Outlook.TaskItem Then
      Set T1 = obj
      Set T2 = F.Items.Add(olTaskItem)

      With T1
        T2.Categories = .Categories
        T2.Companies = .Companies
        T2.ContactNames = .ContactNames
        T2.Subject = .Subject
      End With

      Set Links1 = T1.Links
      Set Links2 = T2.Links

      On Error Resume Next
      For i = 1 To Links1.Count
        Set Link = Links1(i)
        Links2.Add Link.Item
      Next

      'T2.Save
      T2.Display
    End If

    Set T1 = Nothing: Set T2 = Nothing
    Set J1 = Nothing: Set Links1 = Nothing
    Set J2 = Nothing: Set Links2 = Nothing
    Set obj = Nothing

  End If
End Sub
