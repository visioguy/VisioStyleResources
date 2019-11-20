Attribute VB_Name = "MArrowheadShapes"
Option Explicit

'// Module: MArrowheadShapes
'//
'// 2019.11.20
'// Chris Roth
'//
'// A module for automating the creation of arrowhead icon
'// shapes, including creating a vertical stack of shapes,
'// as well as exporting them to pngs of various resolutions.
'//
'// References: Microsoft Scripting Runtime


'// ----- Public Procedures ---------------------------------------------------

Public Sub MakeAllArrowIconShapes(ByRef visShpCaller As Visio.Shape)

  '// Create a vertical stack of indexed icon shapes. Go up
  '// to the minimum index from the calling shape, and down
  '// to the maximum index.
  
  '// Sample ShapeSheet call:
  '// CALLTHIS("MArrowheadShapes.MakeAllArrowIconShapes")
  
  Const minIndex As Integer = 0
  Const maxIndex As Integer = 45
  
  If (IsArrowIconShape(visShpCaller) = False) Then
    Call MsgBox("Shape is not a Line End Icon shape!")
    GoTo Cleanup
  End If
  
  Dim currIndex As Integer
  currIndex = GetArrowIndex(visShpCaller)
  
  Call MUtils.CreateVerticalColumnOfTextIndexShapes(visShpCaller, currIndex, minIndex, maxIndex)
    
Cleanup:
  '
End Sub

Public Sub ExportArrowheadIcons(ByRef visShpCaller As Visio.Shape)

  '// Generate icons files for all Arrowhead Icon shapes
  '// on the current page.

  '// Sample ShapeSheet call:
  '// CALLTHIS("MArrowheadShapes.GenerateArrowheadIcons")

  '// Get the output icon size. A negative size means
  '// the user cancelled or entered nonsense:
  Dim iconSize As Integer
  iconSize = MUtils.GetIconSizeFromUser(1, 1024, 32)
  If (iconSize < 0) Then GoTo ErrorHandler
  
  '// Get the shapes to be exported:
  Dim pg As Visio.Page
  Set pg = visShpCaller.ContainingPage
  
  Dim sel As Visio.Selection
  Set sel = GetShapesByClass(pg, "visguy.visio.ui.lineends.thumbnail")
  If (sel.Count = 0) Then GoTo ErrorHandler
    
  '// Pick a path:
  Dim p As String
  p = MUtils.GetPath_DateTimeAndSuffix(visShpCaller.Document, "Arrowheads")
    
  '// Calculate the dpi:
  Dim w As Double
  w = visShpCaller.CellsU("Width").ResultIU
  
  Dim dpi As Double
  dpi = Int(iconSize / w + 0.5)
  
  '// Notify the user, give them one more chance:
  Dim iconSizeInfo As String
  iconSizeInfo = iconSize & " x " & iconSize
  If (MUtils.UserProceedWithExport(sel.Count, iconSizeInfo, p) = False) Then GoTo ErrorHandler
        
  '// Set the custom export settings:
  Dim xsets As CExportSettings
  Set xsets = New CExportSettings
  Call xsets.SetExportResolutionSettings(visShpCaller.Application, dpi)
          
  '// Create a subdirectory:
  If (MUtils.CreateFolder(p) = False) Then GoTo ErrorHandler
  
  '// Export each shape, and build a new selection of all the shapes:
  Dim arrowIndex As Integer
  Dim fn As String
  Dim shp As Visio.Shape
  For Each shp In sel
    arrowIndex = GetArrowIndex(shp)
    fn = arrowIndex & "_" & iconSize & "x" & iconSize & ".png"
    Call shp.Export(p & fn)
  Next shp
    
  '// Export the whole strip:
  If (sel.Count > 0) Then
    Call sel.Export(p & "_allIcons_" & iconSize & ".png")
  End If
  
  '// Restore the settings:
  Call xsets.RestoreExportResultionSettings
  
  '// Open the folder where the files were exported:
  Dim msg As String
  msg = "Export completed!" & vbCrLf & vbCrLf & "Would you like to view the output directory?"
  If (MsgBox(msg, vbYesNo) = vbYes) Then
    Call Shell("explorer.exe " & p)
  End If
  
  GoTo Cleanup
  
ErrorHandler:
  Call MsgBox("No images will be exported!")
Cleanup:
  Set sel = Nothing
  Set shp = Nothing
  Set pg = Nothing
  Set xsets = Nothing
End Sub

Public Function GetArrowIndex(ByRef visShp As Visio.Shape) As Integer
  GetArrowIndex = 0
  On Error Resume Next
  GetArrowIndex = visShp.Cells("User.Arrow").ResultIU
End Function

Public Function IsArrowIconShape(ByRef visShp As Visio.Shape) As Boolean
  IsArrowIconShape = IsUserClass(visShp, "visguy.visio.ui.lineends.thumbnail")
End Function

