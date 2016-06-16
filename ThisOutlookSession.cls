VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Option Explicit
Private objNS As Outlook.NameSpace
Private WithEvents objNewMailItems As Outlook.Items
Attribute objNewMailItems.VB_VarHelpID = -1
Private WithEvents Items As Outlook.Items
Attribute Items.VB_VarHelpID = -1
 
Private Sub Application_Startup()
  Dim Ns As Outlook.NameSpace
 
  Set Ns = Application.GetNamespace("MAPI")
  Set Items = Ns.GetDefaultFolder(olFolderCalendar).Items
  
  
  
  Dim objMyInbox As Outlook.MAPIFolder
  Set objNS = Application.GetNamespace("MAPI")
  Set objMyInbox = objNS.GetDefaultFolder(olFolderInbox)
  Set objNewMailItems = objMyInbox.Items
  Set objMyInbox = Nothing
  
  
  
  
  
  
End Sub

'http://www.slipstick.com/developer/reminder-all-day-events/
Private Sub Items_ItemAdd(ByVal Item As Object)
 On Error Resume Next
 Dim Appt As Outlook.AppointmentItem

 If TypeOf Item Is Outlook.AppointmentItem Then
  Set Appt = Item

'Checks to see if all day and if it has a reminder set to true
  If Appt.AllDayEvent = True And Appt.ReminderSet = True Then
 
'msgbox block - 3 lines
   If MsgBox("Do you want to remove the reminder?", vbYesNo) = vbNo Then
    Exit Sub
   End If
 
'appt.reminderset block - 2 lines
   Appt.ReminderSet = False
   Appt.Save
 
  End If
 End If
End Sub



'Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
'Dim objEmail As Outlook.MailItem
'Dim strIDs() As String
'Dim intX As Integer
'strIDs = Split(EntryIDCollection, ",")
'For intX = 0 To UBound(strIDs)
'Set objEmail = objNS.GetItemFromID(strIDs(intX))
'Debug.Print "Message subject: " & objEmail.Subject
'Debug.Print "Message sender:" & objEmail.SenderName & " (" & objEmail.SenderEmailAddress & ")"
'Next
'Set objEmail = Nothing
'End Sub




Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
 On Error GoTo ErrorHandler
 Dim strSubject As String
 Dim strUserForm As String
 Dim objRecip As Recipient
 Dim strMsg As String
 Dim Res As Integer
 Dim strBcc As String
 Dim Project_Selected As String 'Selected project in Userform
 
 '------------------------------------------------------------
 '- Autoforward to myself
 '------------------------------------------------------------
 strBcc = "raag@lundbeck.com"
 Set objRecip = Item.Recipients.Add(strBcc)
 objRecip.Type = olBCC
 If Not objRecip.Resolve Then
  strMsg = "Could not resolve the Bcc recipient. " & _
  "Do you want still to send the message?"
  Res = MsgBox(strMsg, vbYesNo + vbDefaultButton1, _
  "Could Not Resolve Bcc Recipient")
  If Res = vbNo Then
   Cancel = True
  End If
 End If
 Set objRecip = Nothing
  
 Exit Sub
 
ErrorHandler:
    MsgBox "HWHAP: " & vbNewLine & Err.Description
    Set objRecip = Nothing
    Unload UserForm1
End Sub








