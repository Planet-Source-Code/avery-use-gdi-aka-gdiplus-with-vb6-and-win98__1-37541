Attribute VB_Name = "mdlWrappers"
Option Explicit

' GDI+ Wrapper By: Dana Seaman
' See more of Dana's wrappers on PSC!
' Search for GDI+ in Visual Basic or go to:
'   http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=38644&lngWId=1
Public Sub DrawGdipSpotLight(ByVal lhdc As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal RadiusX As Long, _
   ByVal RadiusY As Long, _
   ByVal HotX As Long, _
   ByVal HotY As Long, _
   ByVal FocusScale As Single, _
   ByVal StartColor As Long, _
   ByVal EndColor As Long)


   Dim graphics         As Long
   Dim path             As Long
   Dim brush            As Long
   Dim ptl              As POINTL


   ' Initializations
   GdipCreateFromHDC lhdc, graphics
   GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
   GdipCreatePath FillModeWinding, path
   GdipAddPathEllipseI path, x - RadiusX, _
      y - RadiusY, _
      RadiusX + RadiusX, RadiusY + RadiusY
   GdipCreatePathGradientFromPath path, brush
   GdipSetPathGradientCenterColor brush, StartColor
   ptl.x = HotX
   ptl.y = HotY
   GdipSetPathGradientCenterPointI brush, ptl
   GdipSetPathGradientFocusScales brush, FocusScale, FocusScale
   GdipSetPathGradientSurroundColorsWithCount brush, EndColor, 1
   GdipFillPath graphics, brush, path


Cleanup:
   GdipDeleteBrush brush
   GdipDeletePath path
   GdipDeleteGraphics graphics
End Sub


