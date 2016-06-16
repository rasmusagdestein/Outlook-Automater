Attribute VB_Name = "Test_Module"
'Sub Demo_testt()
'    On Error GoTo ErrorHandler
'
' Dim obj As Object
' Dim Sel As Outlook.Selection
' Dim strCategories As String
' Dim strSubject As String
' Dim strCurrentProject As String
' Dim intPosStart As Integer
' Dim intPosSlut As Integer
'
' Dim i As Long
' Dim mySelections() As String
'
' Set Sel = Application.ActiveExplorer.Selection
'  If Sel.Count Then
'   Set obj = Sel(1)
'   If TypeOf obj Is Outlook.JournalItem Then
'    MsgBox ("JournalItem")
'   ElseIf TypeOf obj Is Outlook.MailItem Then
'    strSubject = obj.Subject
''
'
'
'  End If
' End If
' Exit Sub
'
'ErrorHandler:
'    MsgBox "HWHAP: " & vbNewLine & Err.Description
'
'End Sub


Sub Demo_testt()

Dim objNS As Outlook.NameSpace: Set objNS = GetNamespace("MAPI")
Dim olFolder As Outlook.MAPIFolder
Set olFolder = objNS.GetDefaultFolder(olFolderInbox)
Dim Item As Object

For Each Item In olFolder.Items
    If TypeOf Item Is Outlook.MailItem Then
        Dim oMail As Outlook.MailItem: Set oMail = Item
        
        If InStr(oMail.Categories, "[{S}Afventer{/S}{ST}{/ST}]") <> 0 Then
        
       Debug.Print oMail.To
       
        
        End If
    End If
Next

End Sub

Sub InsertText()
    Dim myText As String
    myText = "Hello world"

    Dim NewMail As MailItem, oInspector As Inspector
    Set oInspector = Application.ActiveInspector
    If oInspector Is Nothing Then
        MsgBox "No active inspector"
    Else
        Set NewMail = oInspector.CurrentItem
        If NewMail.Sent Then
            MsgBox "This is not an editable email"
        Else
          NewMail.Subject = NewMail.Subject & myText
        
            'If oInspector.IsWordMail Then
            '    ' Hurray. We can use the rich Word object model, with access
            '    ' the caret and everything.
            '    Dim oDoc As Object, oWrdApp As Object, oSelection As Object
            '    Set oDoc = oInspector.WordEditor
            '    Set oWrdApp = oDoc.Application
            '    Set oSelection = oWrdApp.Selection
            '    oSelection.InsertAfter myText
            '    oSelection.Collapse 0
            '    Set oSelection = Nothing
            '    Set oWrdApp = Nothing
            '    Set oDoc = Nothing
            'Else
            '    ' No object model to work with. Must manipulate raw text.
            '    Select Case NewMail.BodyFormat
            '        Case olFormatPlain, olFormatRichText, olFormatUnspecified
            '            NewMail.Body = NewMail.Body & myText
            '        Case olFormatHTML
            '            NewMail.HTMLBody = NewMail.HTMLBody & "<p>" & myText & "</p>"
            '    End Select
            'End If
        End If
    End If
End Sub
