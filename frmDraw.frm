VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDraw 
   BackColor       =   &H80000005&
   Caption         =   "Freehand Curves"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   2100
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   383
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton cmdAlpha 
         Caption         =   "Alpha"
         Height          =   315
         Left            =   4440
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CheckBox chkAA 
         Caption         =   "Anti-Alias"
         Height          =   315
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Toggle Anti-Aliasing Effects"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdSize 
         Caption         =   "Size"
         Height          =   315
         Left            =   1860
         TabIndex        =   2
         ToolTipText     =   "Change Size"
         Top             =   120
         Width           =   1275
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "Change Color"
         Height          =   315
         Left            =   480
         TabIndex        =   1
         ToolTipText     =   "Change Color"
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label lblColor 
         Height          =   315
         Left            =   60
         TabIndex        =   4
         Top             =   120
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Avery P. - 9/4/2002
' NOTES:
'   - This form *badly* demonstrates how to get VB to handle GDI+ freehand shape drawing, specifically curves.
'     I say badly because there are many optimizations that are not implemented, such as figuring out a way not
'     to have to redraw the entire curve for each added point. Redrawing the entire curve for each new point will
'     result in a major loss of performance the larger the curve gets. There are also some other things wrong, which you'll hopefully spot.
'   - The Curve type/structure can also be expanded and renamed for use with basically any shape and could hold more
'     specific data.
'
' WARNINGS:
'   - This form assumes that the GDI+ Startup/Shutdown was called previously.
'

Option Explicit
Option Base 1  ' We'll make this example a bit more standardized

' Structure for storing shape/curve data
Private Type Curve
   pt() As POINTF ' More precision with singles, but faster processing with longs
   count As Long  ' Number of valid points in the array
   color As Long  ' Color of this curve/shape. Stored in GDI+ format for hopefully obvious reasons.
   size As Long   ' Size of this curve/shape pen width
   AA As Boolean  ' Use Anti-Alias? Default = False
   'type As ShapeType      ' There is no shape type enum; this is for you to have fun with!
End Type

Dim bDrawing As Boolean    ' Drawing a curve? flag.
Dim Curves() As Curve      ' Array of shapes/curves.
Dim lngCurCurve As Long    ' Array index of the current curve; set in the MouseDown and used when adding points.
Dim lngCurColor As Long    ' WARNING: This color is in GDI+ format. See ChangeColor() for conversion functions.
Dim lngCurVBColor As Long  ' WARNING: This color is in VB/Standard RGB format. See ChangeColor() for conversion functions.
Dim lngCurSize As Long     ' Current pen width, in pixels.
Dim bteCurAlpha As Byte    ' Current alpha blending value.

Private Sub cmdAlpha_Click()
   Dim alpha As Long, s As String
   
   s = InputBox("Enter an alpha transparency value." & vbCrLf & "   0 = Totally transparent" & vbCrLf & "   255 = Totally opaque" & vbCrLf & "WARNING: If you change the value, the current method of redrawing the entire curve for each new point will conflict with your chosen value. You can optimize the drawing or watch the neat side-effect that results - it is very nice, especially if you choose a very low alpha value.", "Change Alpha-Blending Value", bteCurAlpha)
   If Len(s) Then
      alpha = Val(s)
      If alpha < 0 Then alpha = 0
      If alpha > 255 Then alpha = 255
      ChangeAlpha alpha
   End If
End Sub

Private Sub cmdColor_Click()
On Error GoTo errColor
   dlgColor.color = lngCurVBColor
   dlgColor.flags = cdlCCFullOpen Or cdlCCRGBInit
   dlgColor.ShowColor
   ChangeColor dlgColor.color
Exit Sub
errColor:
   Debug.Print Err.Number & "; " & Err.description
End Sub

Private Sub cmdSize_Click()
   Dim size As Long, s As String
   
   s = InputBox("Enter a new size:", "Change Drawing Width", lngCurSize)
   If Len(s) Then
      size = Val(s)
      If size <= 0 Then size = 1
      ChangeSize size
   End If
End Sub

Private Sub Form_Load()
   ' Initialize some vars
   bDrawing = False
   Me.AutoRedraw = True ' Speeds things up a bit
   lngCurCurve = 0
   
   ' Set some more defaults
   ChangeColor Aquamarine, True
   ChangeSize 1
   ChangeAlpha 255
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   ' Clear all allocated arrays
   Dim i As Long
   For i = 1 To UBound(Curves)
      Erase Curves(i).pt
   Next
   Erase Curves
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton And bDrawing = False Then
      bDrawing = True

      ' Allocate room for another shape/curve
      lngCurCurve = lngCurCurve + 1
      ReDim Preserve Curves(lngCurCurve)
      
      ' Save the current drawing settings
      Curves(lngCurCurve).color = ColorSetAlpha(lngCurColor, bteCurAlpha)
      Curves(lngCurCurve).size = lngCurSize
      Curves(lngCurCurve).AA = CBool(chkAA.value)
      
      ' Initialize the first point
      AddPoint Curves(lngCurCurve), x, y
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If bDrawing Then
      AddPoint Curves(lngCurCurve), x, y
   End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton And bDrawing Then
      bDrawing = False
      AddPoint Curves(lngCurCurve), x, y
   End If
End Sub

' Called when there is no AutoRedraw
Private Sub Form_Paint()
   DrawCurves
End Sub



' Add a new point to a curve
' We'll double the number of points should we run out. This should speed things up a bit.
Private Sub AddPoint(aCurve As Curve, ByVal x As Single, ByVal y As Single)
   ' Prevent adding the exact same coord twice in a row - it happens and is wasteful!
   If aCurve.count > 0 Then   ' Must ensure we have a point to check against!
      With aCurve.pt(aCurve.count)
         If .x = x And .y = y Then Exit Sub
      End With
   End If

   ' It is not a duplicate entry...
   ' Increase the count
   aCurve.count = aCurve.count + 1
   
   ' Allocate extra memory as needed
   If aCurve.count <= 1 Then
      ReDim aCurve.pt(1)
   ElseIf aCurve.count > UBound(aCurve.pt) Then
      ' I'm trying to cut down on the point array redimming, as it can negatively impact performance
      ' and the points will be flowing in like water! Yet you must keep the wasted memory low...
      ReDim Preserve aCurve.pt(UBound(aCurve.pt) + 10)
   End If
   
   ' Save the point in the array
   With aCurve.pt(aCurve.count)
      .x = x
      .y = y
   End With

   ' Draw the curve
   DrawCurve aCurve
End Sub

Private Sub DrawCurves()
   Dim i As Long
   Dim graphics As Long, pen As Long
   
   ' Exit if there are no curves to draw
   If lngCurCurve <= 0 Then Exit Sub
   
   ' Initialize GDI+ drawing
   Call GdipCreateFromHDC(Me.hdc, graphics)
   
   For i = 1 To lngCurCurve
      ' To AntiAlias or not to AntiAlias?
      If Curves(i).AA Then
         Call GdipSetSmoothingMode(graphics, SmoothingModeAntiAlias)
      Else
         Call GdipSetSmoothingMode(graphics, SmoothingModeNone)
      End If

      ' Each shape/curve has its own properties, so we must create a pen for each!
      Call GdipCreatePen1(Curves(i).color, Curves(i).size, UnitPixel, pen)
      Call GdipDrawCurve(graphics, pen, Curves(i).pt(1), Curves(i).count)
      Call GdipDeletePen(pen) ' Cleanup
   Next
   
   ' Cleanup
   GdipDeleteGraphics graphics
End Sub

Private Sub DrawCurve(aCurve As Curve)
   Dim graphics As Long, pen As Long
   
   
   ' Initialize GDI+ drawing
   Call GdipCreateFromHDC(Me.hdc, graphics)
   Call GdipCreatePen1(aCurve.color, aCurve.size, UnitPixel, pen)
   
   ' To AntiAlias or not to AntiAlias?
   If aCurve.AA Then
      Call GdipSetSmoothingMode(graphics, SmoothingModeAntiAlias)
   Else
      Call GdipSetSmoothingMode(graphics, SmoothingModeNone)
   End If

   ' Draw
   Call GdipDrawCurve(graphics, pen, aCurve.pt(1), aCurve.count)
   
   ' Force a refresh if AutoRedrawing
   ' NOTE: This is a workaround for the bug that is published in Q189736
   If Me.AutoRedraw Then Me.Refresh
   
   ' Cleanup
   GdipDeletePen pen
   GdipDeleteGraphics graphics
End Sub

Private Sub ChangeSize(ByVal lSize As Long)
   lngCurSize = lSize
   cmdSize.Caption = "Size: " & lSize & " pixel(s)"
End Sub

Private Sub ChangeColor(ByVal lColor As Long, Optional ByVal bGDIPColor As Boolean)
   If bGDIPColor = True Then
      lngCurColor = lColor
      lngCurVBColor = GetRGB_GDIP2VB(lColor)
   Else
      lngCurColor = GetRGB_VB2GDIP(lColor)
      lngCurVBColor = lColor
   End If
   lblColor.BackColor = lngCurVBColor
End Sub

Private Sub ChangeAlpha(ByVal alpha As Byte)
   bteCurAlpha = alpha
   cmdAlpha.Caption = "Alpha: " & alpha
End Sub
