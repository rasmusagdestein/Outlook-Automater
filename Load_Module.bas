Attribute VB_Name = "Load_Module"
Sub Load_Script()
 On Error GoTo ErrorHandler
 
 Dim objNameSpace As NameSpace
 Dim objCategory As Category
 Dim strOutput As String
 
 Dim intAreaCount As Integer
 Dim intProjectCount As Integer
 Dim intManufacturerCount As Integer
 Dim intStatusCount As Integer
 
 Dim preProjects() As String
 Dim preAreas() As String
 Dim preManufacturers() As String
 Dim preStatus() As String
 
 Dim result As Boolean
 Dim intPosStart As Integer
 Dim intPosSlut As Integer
 Dim i As Integer
 Dim j As Integer
 Dim k As Integer
 Dim l As Integer
 
 ' Obtain a NameSpace object reference.
 Set objNameSpace = Application.GetNamespace("MAPI")

 ' Check if the Categories collection for the Namespace
 ' contains one or more Category objects.
 If objNameSpace.Categories.Count > 0 Then
 
  ' Amount of projects and areas. ------------------------------------------------------
  intAreaCount = 0
  intProjectCount = 0
  intManufacturerCount = 0
  intStatusCount = 0
  For Each objCategory In objNameSpace.Categories
   If InStr(objCategory.Name, "[{P") <> 0 Then intProjectCount = intProjectCount + 1
   If InStr(objCategory.Name, "[{L") <> 0 Then intAreaCount = intAreaCount + 1
   If InStr(objCategory.Name, "[{M") <> 0 Then intManufacturerCount = intManufacturerCount + 1
   If InStr(objCategory.Name, "[{S") <> 0 Then intStatusCount = intStatusCount + 1
  Next
  '--------------------------------------------------------------------------------------
 
  ' Parse areas and projects into global arrays -----------------------------------------
  If (intAreaCount > 0) And (intProjectCount > 0) And (intManufacturerCount > 0) And (intStatusCount > 0) Then
   ReDim preAreas(0 To intAreaCount - 1)
   ReDim preProjects(0 To intProjectCount - 1)
   ReDim preManufacturers(0 To intManufacturerCount - 1)
   ReDim preStatus(0 To intStatusCount - 1)
   
  
   i = 0
   j = 0
   k = 0
   l = 0
  
   For Each objCategory In objNameSpace.Categories
    ' Projects
    If InStr(objCategory.Name, "[{P") <> 0 Then
     preProjects(i) = objCategory.Name
     i = i + 1
    End If
    ' Areas
    If InStr(objCategory.Name, "[{L") <> 0 Then
     preAreas(j) = objCategory.Name
     j = j + 1
    End If
    ' Manufacturers
    If InStr(objCategory.Name, "[{M") <> 0 Then
     preManufacturers(k) = objCategory.Name
     k = k + 1
    End If
    ' Status
    If InStr(objCategory.Name, "[{S") <> 0 Then
     preStatus(l) = objCategory.Name
     l = l + 1
    End If

   Next
  
   'Sort arrays
   result = QSortInPlace(preProjects)
   result = QSortInPlace(preAreas)
   result = QSortInPlace(preManufacturers)
   result = QSortInPlace(preStatus)
   
 
 
 
    'Parse Projects
   ReDim NewProjects(0 To intProjectCount, 0 To 3)
  
   NewProjects(0, 0) = "None"
   NewProjects(0, 1) = "None"
   NewProjects(0, 2) = ""
   NewProjects(0, 3) = ""
   
   
   For i = 1 To intProjectCount
    intPosStart = InStr(preProjects(i - 1), "{P}")
    intPosSlut = InStr(preProjects(i - 1), "{/P}")
    NewProjects(i, 0) = Mid(preProjects(i - 1), intPosStart + 3, intPosSlut - intPosStart - 3)
  
    intPosStart = InStr(preProjects(i - 1), "{T}")
    intPosSlut = InStr(preProjects(i - 1), "{/T}")
    NewProjects(i, 1) = Mid(preProjects(i - 1), intPosStart + 3, intPosSlut - intPosStart - 3)
  
    intPosStart = InStr(preProjects(i - 1), "{A}")
    intPosSlut = InStr(preProjects(i - 1), "{/A}")
    NewProjects(i, 2) = Mid(preProjects(i - 1), intPosStart + 3, intPosSlut - intPosStart - 3)
    
    NewProjects(i, 3) = preProjects(i - 1)
   Next i
 
 
 
   'Parse Areas
   ReDim NewAreas(0 To intAreaCount - 1, 0 To 2)
   For i = 0 To intAreaCount - 1
    intPosStart = InStr(preAreas(i), "{L}")
    intPosSlut = InStr(preAreas(i), "{/L}")
    NewAreas(i, 0) = Mid(preAreas(i), intPosStart + 3, intPosSlut - intPosStart - 3)
    
    intPosStart = InStr(preAreas(i), "{LT}")
    intPosSlut = InStr(preAreas(i), "{/LT}")
    NewAreas(i, 1) = Mid(preAreas(i), intPosStart + 4, intPosSlut - intPosStart - 4)

    NewAreas(i, 2) = preAreas(i)
   Next i
   
   'Parse Manufacturers
   ReDim Manufacturers(0 To intManufacturerCount - 1, 0 To 2)
   For i = 0 To intManufacturerCount - 1
    intPosStart = InStr(preManufacturers(i), "{M}")
    intPosSlut = InStr(preManufacturers(i), "{/M}")
    Manufacturers(i, 0) = Mid(preManufacturers(i), intPosStart + 3, intPosSlut - intPosStart - 3)
   
    intPosStart = InStr(preManufacturers(i), "{MT}")
    intPosSlut = InStr(preManufacturers(i), "{/MT}")
    Manufacturers(i, 1) = Mid(preManufacturers(i), intPosStart + 4, intPosSlut - intPosStart - 4)
    
    Manufacturers(i, 2) = preManufacturers(i)

   Next i
   
   'Parse Status
   ReDim Status(0 To intStatusCount - 1, 0 To 2)
   For i = 0 To intStatusCount - 1
    intPosStart = InStr(preStatus(i), "{S}")
    intPosSlut = InStr(preStatus(i), "{/S}")
    Status(i, 0) = Mid(preStatus(i), intPosStart + 3, intPosSlut - intPosStart - 3)
  
    intPosStart = InStr(preStatus(i), "{ST}")
    intPosSlut = InStr(preStatus(i), "{/ST}")
    Status(i, 1) = Mid(preStatus(i), intPosStart + 4, intPosSlut - intPosStart - 4)
    
    Status(i, 2) = preStatus(i)

   Next i
   
   
  End If
 End If
 

 
 ' Clean up.
 Set objCategory = Nothing
 Set objNameSpace = Nothing

 
 Exit Sub

ErrorHandler:
 MsgBox "HWHAP: " & vbNewLine & Err.Description
    



End Sub


