VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExportSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// Class: CExportSettings
'//
'// 2019.11.20
'// Chris Roth
'//
'// A class for encapsulating the nitpicky details of saving,
'// setting, then restoring raster-image export settings.

Private m_visAppSets As Visio.ApplicationSettings

'// Variables for storing the initial export settings:
Private m_xRes0 As Visio.VisRasterExportResolution
Private m_xWidth0 As Double
Private m_xHeight0 As Double
Private m_xUnits0 As Visio.VisRasterExportResolutionUnits

'// ----- Constructors, Destructors -------------------------------------------
Private Sub Class_Initialize()
  '
End Sub
Private Sub Class_Terminate()
  Set m_visAppSets = Nothing
End Sub

'// ----- Public Procedures ---------------------------------------------------
Public Sub SetExportResolutionSettings( _
      ByRef visApp As Visio.Application, _
      ByVal dotsPerInch As Integer)
    
  '// Save what is there, first:
  Set m_visAppSets = visApp.Settings
  Call m_visAppSets.GetRasterExportResolution(m_xRes0, m_xWidth0, m_xHeight0, m_xUnits0)
   
  '// Note: Visio provides independent horizontal and vertical resolution,
  '// but who ever needs that for exporting? This class simplifies that by
  '// only offering one dpi parameter.
   
  '// Now set to custom, pixels-per-inch:
  Call m_visAppSets.SetRasterExportResolution( _
      Visio.VisRasterExportSize.visRasterFitToCustomSize, _
      dotsPerInch, dotsPerInch, _
      Visio.VisRasterExportResolutionUnits.visRasterPixelsPerInch)
      
End Sub

Public Sub RestoreExportResultionSettings()

  If (m_visAppSets Is Nothing) Then Exit Sub

  '// Restore the export settings to whatever was saved last:
  Call m_visAppSets.SetRasterExportResolution( _
        m_xRes0, _
        m_xWidth0, m_xHeight0, _
        m_xUnits0)
        
End Sub


