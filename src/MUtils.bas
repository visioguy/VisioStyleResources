Attribute VB_Name = "MUtils"
Option Explicit

'// Module: MUtils
'//
'// 2019.11.20
'// Chris Roth
'//
'// A module for shared functions, such as IO and file operations,
'// generic Visio procedures, and user alerts and prompts.

Public Function CreateFolder(ByVal path As String) As Boolean

  On Error GoTo ErrorHandler
  
  CreateFolder = False

  '// Create a subdirectory:
  Dim fso As Scripting.FileSystemObject
  Set fso = New Scripting.FileSystemObject
  Call fso.CreateFolder(path)
  
  CreateFolder = True
  
  GoTo Cleanup
  
ErrorHandler:
  Debug.Print "Error in MUtils.CreateFolder:" & vbCrLf & Error$
Cleanup:
  Set fso = Nothing
End Function
Public Sub CreateVerticalColumnOfTextIndexShapes( _
      ByRef visShp As Visio.Shape, _
      ByVal currIndex As Integer, _
      ByVal minIndex As Integer, _
      ByVal maxIndex As Integer)

  '// This proc creates a tightly-stacked vertical column of shapes.
  '//
  '// - The shapes have an index that is indicated by the text of the shape.
  '// - The proc handles the creation of shapes that come before and after
  '//   visShp. This is determined by currIndex being compared to the minIndex
  '//   and to maxIndex.
  '//
  '// For example:
  '// If the current index for visShp is 5 and we want to generate a stack
  '// of shapes with indices from 2 to 17, the proc will create shapes 2-4
  '// immediately above visShp, then it will add shapes 6-17 immediately
  '// below visShp.
  '//
  '// For each index, a duplicate of visShp will be created, properly-positioned,
  '// and the text for the duplicate will be set to the new indes.

  Dim px As Double, py As Double
  Dim h As Double
  Dim i As Integer, di As Integer
  Dim shpCopy As Visio.Shape
  
  px = visShp.CellsU("PinX").ResultIU
  py = visShp.CellsU("PinY").ResultIU
  h = visShp.CellsU("Height").ResultIU
     
  '// Go up/backwards:
  di = 1
  For i = currIndex - 1 To minIndex Step -1
    Set shpCopy = visShp.Duplicate
    shpCopy.CellsU("PinX").ResultIU = px
    shpCopy.CellsU("PinY").ResultIU = py + h * di
    shpCopy.Text = i
    di = di + 1
  Next i
  
  '// Go down/forwards:
  di = -1
  For i = currIndex To maxIndex
    Set shpCopy = visShp.Duplicate
    shpCopy.CellsU("PinX").ResultIU = px
    shpCopy.CellsU("PinY").ResultIU = py + h * di
    shpCopy.Text = i
    di = di - 1
  Next i
  
Cleanup:
  Set shpCopy = Nothing
End Sub

Public Function GetIconSizeFromUser( _
      ByVal minSize As Integer, ByVal maxSize As Integer, _
      ByVal defSize As Integer) As Integer
  
  '// A function that gets an icon size from the user via the InputBox function.
  
  GetIconSizeFromUser = -1
  
  Dim s As String
  s = VBA.InputBox("Icon Size (" & minSize & "-" & maxSize & "):", "Output Icon Size", defSize)
  
  '// Cancelled?
  If (Len(s) = 0) Then GoTo Cleanup
  
  '// Box the icon size:
  Dim iconSize As Integer
  iconSize = val(s)
  If (iconSize = 0) Then GoTo Cleanup
  
  If (iconSize < 0) Then iconSize = 1
  If (iconSize < 1) Then iconSize = Abs(minSize)
  If (iconSize > 1024) Then iconSize = Abs(maxSize)
  
  GetIconSizeFromUser = iconSize
  
Cleanup:
  '
End Function
Public Function GetPath_DateTimeAndSuffix( _
      ByRef visDoc As Visio.Document, _
      ByVal suffix As String) As String
  
  '// Returns the path to a subfolder in the directory where
  '// visDoc is saved. If visDoc is not saved, an untrapped
  '// error will occur. The current date will start the folder
  '// name using year, month, day, hours, minutes, second. The
  '// suffix will end the folder name.
  '//
  '// The format of the path will look like this:
  '//
  '// D:\Work\Projects\Visio\20191120_153522 LinePatterns\
  '//
  '// where "LinePatterns" is the suffix parameter.
  
  Dim p As String
  p = visDoc.path '//...eg: D:\Work\Projects\Visio Resources\LineEnds (Arrowheads)\
  p = p & Format(VBA.Now, "YYYYMMDD_HHmmss") & " " & suffix & "\"
  
  GetPath_DateTimeAndSuffix = p
  
End Function
Public Function GetShapesByClass(ByRef visPg As Visio.Page, ByVal userClassValue As String) As Visio.Selection

  Dim sel As Visio.Selection
  Set sel = visPg.CreateSelection(Visio.VisSelectionTypes.visSelTypeEmpty)
  
  Dim shp As Visio.Shape
  For Each shp In visPg.Shapes
    If IsUserClass(shp, userClassValue) Then
      Call sel.Select(shp, Visio.VisSelectArgs.visSelect)
    End If
  Next shp
     
  Set GetShapesByClass = sel
  
Cleanup:
  Set shp = Nothing
  Set sel = Nothing
End Function
Public Function IsUserClass(ByRef visShp As Visio.Shape, ByVal userClassValue As String) As Boolean

  '// Shapes of interest can be identified by adding the the user-defined
  '// cell "User.Class" to a shape, then setting a unique string value there.
  '// I like to use User.Class="this.that.something" instead of depending
  '// on master names, since master names can end up with "dot-number" suffixes
  '// when newer versions are added to a document (eg: "PC", "PC.36", "PC.129")
  '//
  '// Also, User.Class makes it possible to have different shapes that are
  '// functionally the same be quickly identified by code. It may be that shapes
  '// have different graphics, but the same set of Shape Data fields or other
  '// behaviors.
  
  IsUserClass = False
  
  If (visShp.CellExists("User.Class", Visio.VisExistsFlags.visExistsAnywhere)) Then
  
    Dim val As String
    val = visShp.Cells("User.Class").ResultStr(Visio.VisUnitCodes.visUnitsString)
    
    IsUserClass = (val = userClassValue)
  
  End If

End Function
Public Function UserProceedWithExport( _
      ByVal shapeCount As Integer, _
      ByVal iconSizeInfo As String, _
      ByVal directoryPath As String) As Boolean

  '// Another Yes/No message box alert for the user that sums up
  '// how many shapes have been identified for export, along with
  '// the export destination and the icon size.
  
  UserProceedWithExport = False

  '// Notify the user, give them one more chance:
  Dim msg As String
  msg = "Pre-export summary:" & vbCrLf & vbCrLf
  msg = msg & "Shapes: " & vbTab & vbTab & shapeCount & vbCrLf
  msg = msg & "Image size: " & vbTab & iconSizeInfo & vbCrLf & vbCrLf
  msg = msg & "Target export directory: "
  msg = msg & vbCrLf & vbCrLf
  msg = msg & "'" & directoryPath & "'"
  msg = msg & vbCrLf & vbCrLf
  msg = msg & "Would you like to proceed?"
  
  UserProceedWithExport = (MsgBox(msg, vbYesNo) = vbYes)
  
End Function
