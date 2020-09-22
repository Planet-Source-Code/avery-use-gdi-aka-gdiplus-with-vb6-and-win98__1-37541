VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "GDI+ Test"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7815
   LinkTopic       =   "GDIPlus Test"
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picGrapes 
      Height          =   915
      Left            =   3660
      ScaleHeight     =   855
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   2580
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFree 
         Caption         =   "FreeHand Test"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRedraw 
      Caption         =   "&Redraw/Start Demo"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Avery P. - 7/30/2002
' Examples are from the GDI+ portion of the Platform SDK
' NOTE: If you think this code is a bit sloppy, you are probably right!



' Used for my very simple GIF demo. You should use the Multimedia Timer Functions for production apps.
Private Declare Function GetTickCount Lib "kernel32" () As Long

' And this with the PixelFormat constant to get the bpp
' The trailing & is needed as a workaround for the hex value, else VB will treat it as an integer!?!
Private Const PixelFormatBPPMask As Long = &HFF00&

Dim token As Long ' Needed to close GDI+
Dim bHitTesting As Boolean    ' Used for hit test demo


Private Sub Form_Load()
   ' Load the GDI+ Dll
   Dim GpInput As GdiplusStartupInput
   GpInput.GdiplusVersion = 1
   If GdiplusStartup(token, GpInput) <> Ok Then
      MsgBox "Error loading GDI+!", vbCritical
      Unload Me
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If bHitTesting Then Call DrawHitTest(x, y)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Unload the GDI+ Dll
   Call GdiplusShutdown(token)
End Sub

Private Sub mnuFree_Click()
On Error Resume Next
   Load frmDraw
   frmDraw.Show vbModal
   Set frmDraw = Nothing
End Sub

Private Sub mnuRedraw_Click()
   Cls   ' Clear the window; flicker...yay!

   ' Uncomment one to see it's demo
   'Call DrawCurves
   'Call DrawScaling
   'Call DrawSkewed
   'Call DrawTexturedLine
   'Call DrawLineCaps
   'Call DrawCustomDashed
   'Call DrawSolidShape
   'Call DrawHatchShape
   'Call DrawThumbnail     ' You'll need to change the image filename
   'Call DrawHorizGradient
   'Call DrawDiagGradient
   'Call DrawPathGradient
   'Call BMPtoPNG
   'Call BMPtoJPEG
   'Call BMPtoJPEG_Params
   'Call DrawCachedBitmap  ' WARNING: Running this with AutoRedraw = True is not a good idea!
   Call DrawAlphaLines
   'Call DrawColorMatrix
   'Call DrawAlphaPixels
   'Call DrawAntiAliasText
   'Call DrawFormatText
   'Call DrawRotated
   'Call DrawClippedText
   'Call DrawClippedImage
   'Call DrawHitTest
   Call DrawBezier
   'Call SaveScaling
   'Call SaveStream        ' You'll need to load a typelib with IStorage and IStream in them.
                           ' I used this one: http://www.vbbyjc.com/typelibs/IStorage.tlb
                           ' You could also declare the APIs for them if you like.
   'Call GetAllProperties
   'Call ChangeTitleProperty
   'Call RetrievePathData
   'Call PlayGIF
   'Call CropImage
   'Call SaveGrayscale
   'Call DrawGrid          ' Watch a possible GDI+ flaw in action
   'Call DrawTextBackColor
   'Call LockBitsWriteUIB
   'Call BMPtoGIF
   'Call BMPtoGIF_Transparency
   

   '''''''''''''''''''''''''''''''''''''''''''
   ' Wrapper Demos
   '''''''''''''''''''''''''''''''''''''''''''
   '  Circular Example:
   'DrawGdipSpotLight hdc, 50, 150, 25, 25, 50, 150, 0, Yellow, Yellow And 536870912
   '  Elliptical Example:
   'DrawGdipSpotLight hdc, 100, 250, 100, 50, 50, 250, 0.05, White, Black And &H80000000
End Sub

Private Sub mnuExit_Click()
   Unload Me
End Sub


'======================================================================
' THE SAMPLES!
'======================================================================
Private Sub DrawLineCaps()
   Dim graphics As Long, pen As Long
   
   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics)     ' Initialize the graphics class - required for all drawing
   Call GdipCreatePen1(Blue, 8, UnitPixel, pen)
   
   ' Set the start and end caps
   Call GdipSetPenStartCap(pen, LineCapArrowAnchor)
   Call GdipSetPenEndCap(pen, LineCapRoundAnchor)
   
   ' Draw the line
   Call GdipDrawLineI(graphics, pen, 20, 175, 300, 175)
   
   ' Cleanup
   Call GdipDeletePen(pen)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawCustomDashed()
   Dim graphics As Long, pen As Long
   Dim dashValues(1 To 4) As Single
   
   ' Set the dash intervals
   ' The dashes are in an on/off pattern and continually repeat for the length of the line
   dashValues(1) = 5    ' Show 5 * penwidth
   dashValues(2) = 2    ' Hide 2 * penwidth
   dashValues(3) = 15   ' Show 15 * penwidth
   dashValues(4) = 4    ' Hide 4 * pendwith

   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics)     ' Initialize the graphics class - required for all drawing
   Call GdipCreatePen1(Black, 4, UnitPixel, pen)
   
   ' Set the dash pattern
   Call GdipSetPenDashArray(pen, dashValues(1), 4)
   
   ' Draw the line
   Call GdipDrawLineI(graphics, pen, 5, 5, 405, 5)
   
   ' Cleanup
   Call GdipDeletePen(pen)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawTexturedLine()
   Dim graphics As Long, img As Long, pen As Long, tBrush As Long
   Dim lngHeight As Long, lngWidth As Long

   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics)     ' Initialize the graphics class - required for all drawing
   Call GdipLoadImageFromFile(StrConv(App.path & "\Texture.bmp", vbUnicode), img)
   Call GdipCreateTexture(img, WrapModeTile, tBrush) ' Create a textured brush
   Call GdipCreatePen2(tBrush, 30, UnitPixel, pen)  ' Create a pen to draw with

   ' Get the image height and width
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)

   Call GdipDrawImageRect(graphics, img, 0, 0, lngWidth, lngHeight)
   Call GdipDrawEllipseI(graphics, pen, 100, 20, 200, 100)
   
   ' Cleanup
   Call GdipDeletePen(pen)
   Call GdipDeleteBrush(tBrush)
   Call GdipDisposeImage(img)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawSolidShape()
   Dim graphics As Long, brush As Long
   

   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipCreateSolidFill(DeepPink, brush)     ' Create the solid colored brush
   
   ' Draw an ellipse
   Call GdipFillEllipseI(graphics, brush, 0, 0, 100, 60)

   ' Cleanup
   Call GdipDeleteBrush(brush)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawHatchShape()
   Dim graphics As Long, brush As Long


   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipCreateHatchBrush(HatchStyleDottedGrid, Black, DeepPink, brush)      ' Create the pattern brush
   
   ' Draw an ellipse
   Call GdipFillEllipseI(graphics, brush, 0, 0, 100, 60)

   ' Cleanup
   Call GdipDeleteBrush(brush)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawCurves()
   Dim graphics As Long, pen As Long
   Dim Points(1 To 5) As POINTL
   

   ' Random values (From the SDK C++ sample)
   Points(1).x = 0
   Points(1).y = 100
   Points(2).x = 50
   Points(2).y = 80
   Points(3).x = 100
   Points(3).y = 20
   Points(4).x = 150
   Points(4).y = 80
   Points(5).x = 200
   Points(5).y = 100

   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics)     ' Initialize the graphics class - required for all drawing
   Call GdipCreatePen1(Linen, 10, UnitPixel, pen)  ' Create a pen to draw with

   ' Gamma correction is nice, though slower...
   Call GdipSetCompositingQuality(graphics, CompositingQualityGammaCorrected)
   ' Draw the curve w/ anti-aliasing
   Call GdipSetSmoothingMode(graphics, SmoothingModeAntiAlias)
   Call GdipDrawCurveI(graphics, pen, Points(1), 5)
   
   ' Cleanup
   Call GdipDeletePen(pen)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawScaling()
   ' Why did I store the picture in a hidden PictureBox, you ask?
   ' Well, I'll tell you why: To make the demo simple yet complex
   ' enough for you to figure out how to build on it. (I hate VERY
   ' simple demos, and you should too!) I hope this isn't too simple...
   Dim graphics As Long, img As Long
   Dim lngHeight As Long, lngWidth As Long
   
   
   ' We are going to draw on the form, hence the Me.hDC
   Call GdipCreateFromHDC(Me.hdc, graphics)  ' Initialize the graphics class - required for all drawing
   
   
   ' Load the bitmap file into the Picture box (could also embed it)
   Set picGrapes.Picture = LoadPicture("GrapeBunch.bmp")
   ' WARNING: Make sure the picture box is large enough - Since we will create the image from
   '           the picture box image object, only the viewable area will become the image!
   ' NOTE: You can use the Picture object to get the required handles to bypass this issue.
   picGrapes.AutoSize = True  ' This should do for what we need
   ' Get the image "class" from the PictureBox
   Call GdipCreateBitmapFromHBITMAP(picGrapes.image.Handle, picGrapes.image.hpal, img)
   ' Below is the "cheap" way (via file); good for all supported file type
   ' Could also use GdipCreateBitmapFromFile for this bitmap
   ' Comment out the picture box code above and uncomment this to try it out if you want!
   'Call GdipLoadImageFromFile(StrConv(app.path & "\GrapeBunch.bmp", vbUnicode), img)
   

   ' Get the image height and width
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   
   '**** If you don't pass a width and height when drawing, the image is auto-scaled!! ****
   'Call GdipDrawImage(graphics, img, 10, 10) ' Auto-Scaled

   ' Draw the image with no shrinking or stretching
   Call GdipDrawImageRectI(graphics, img, 10, 10, lngWidth, lngHeight)
   
   ' Shrink the image using low-quality interpolation.
   Call GdipSetInterpolationMode(graphics, InterpolationModeNearestNeighbor)
   Call GdipDrawImageRectRectI(graphics, img, 10, 250, 0.6 * lngWidth, 0.6 * lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel)
   
   ' Shrink the image using medium-quality interpolation.
   Call GdipSetInterpolationMode(graphics, InterpolationModeHighQualityBilinear)
   Call GdipDrawImageRectRectI(graphics, img, 150, 250, 0.6 * lngWidth, 0.6 * lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel)
   
   ' Shrink the image using high-quality interpolation.
   Call GdipSetInterpolationMode(graphics, InterpolationModeHighQualityBicubic)
   Call GdipDrawImageRectRectI(graphics, img, 290, 250, 0.6 * lngWidth, 0.6 * lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel)
   
   ' NOTES: Since we are shrinking the entire image, we could just as well have called
   '        the GdipDrawImageRectI function, which would simplify things - but our goal must
   '        be to make life hellish!
   
   ' Cleanup
   Call GdipDisposeImage(img) ' Delete the image
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawSkewed()
   Dim graphics As Long, img As Long
   Dim destinationPoints(1 To 3) As POINTL
   Dim lngHeight As Long, lngWidth As Long


   ' Set the skewing points in the point array.
   ' destination for upper-left point of original
   destinationPoints(1).x = 200
   destinationPoints(1).y = 20
   ' destination for upper-right point of original
   destinationPoints(2).x = 110
   destinationPoints(2).y = 100
   ' destination for lower-left point of original
   destinationPoints(3).x = 250
   destinationPoints(3).y = 30

   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipLoadImageFromFile(StrConv(App.path & "\Stripes.bmp", vbUnicode), img)   ' Load the image

   ' Get the image height and width
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)

   ' Draw the image unaltered with its upper-left corner at (0, 0).
   Call GdipDrawImageRectI(graphics, img, 0, 0, lngWidth, lngHeight)
   
   ' Draw the image mapped to the parallelogram.
   Call GdipDrawImagePointsI(graphics, img, destinationPoints(1), 3)
   
   ' Cleanup
   Call GdipDisposeImage(img) ' Delete the image
   Call GdipDeleteGraphics(graphics)
End Sub


Private Sub DrawThumbnail()
   Dim graphics As Long, img As Long, imgThumb As Long
   Dim lngHeight As Long, lngWidth As Long


   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipLoadImageFromFile(StrConv(App.path & "\SomeBigImage.bmp", vbUnicode), img)   ' Load the image
   
   ' Create the thumbnail that is 100x100 in size
   Call GdipGetImageThumbnail(img, 100, 100, imgThumb)

   ' Get the image height and width
   ' NOTE: Could also skip this and use the hard-coded/predefined values to save memory and time.
   Call GdipGetImageHeight(imgThumb, lngHeight)
   Call GdipGetImageWidth(imgThumb, lngWidth)

   ' Draw the thumbnail image unaltered
   Call GdipDrawImageRectI(graphics, imgThumb, 10, 10, lngWidth, lngHeight)

   ' Cleanup
   Call GdipDisposeImage(img) ' Delete the image
   Call GdipDisposeImage(imgThumb) ' Delete the thumbnail image
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawHorizGradient()
   Dim graphics As Long, brush As Long, pen As Long
   Dim pt1 As POINTL, pt2 As POINTL


   ' Set the gradient color points
   pt1.x = 0
   pt1.y = 10
   pt2.x = 200
   pt2.y = 10
   
   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   ' Create the gradient brush; we'll use tiling
   Call GdipCreateLineBrushI(pt1, pt2, Red, Blue, WrapModeTile, brush)
   ' Create a pen with the same gradient brush
   Call GdipCreatePen2(brush, 1, UnitPixel, pen)
   
   ' Draw some objects
   Call GdipDrawLine(graphics, pen, 0, 10, 200, 10)
   Call GdipFillEllipse(graphics, brush, 0, 30, 200, 100)
   Call GdipFillRectangle(graphics, brush, 0, 155, 500, 30)

   'Cleanup
   Call GdipDeletePen(pen)
   Call GdipDeleteBrush(brush)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawDiagGradient()
   Dim graphics As Long, brush As Long, pen As Long
   Dim pt1 As POINTL, pt2 As POINTL

   ' Set the gradient color points
   ' pt1 will stay at 0,0
   pt2.x = 200
   pt2.y = 100
   
   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   ' Create the gradient brush; we'll use tiling
   Call GdipCreateLineBrushI(pt1, pt2, LightGreen, LightBlue, WrapModeTile, brush)
   ' Create a pen with the same gradient brush
   Call GdipCreatePen2(brush, 10, UnitPixel, pen)
   
   ' Draw some objects
   Call GdipDrawLineI(graphics, pen, 0, 0, 600, 300)
   Call GdipFillEllipseI(graphics, brush, 10, 100, 200, 100)

   'Cleanup
   Call GdipDeletePen(pen)
   Call GdipDeleteBrush(brush)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawPathGradient()
   Dim graphics As Long, brush As Long, path As Long

   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   ' Create a GraphicsPath object
   Call GdipCreatePath(FillModeAlternate, path)
   
   ' Add an ellipse to the path
   Call GdipAddPathEllipseI(path, 0, 0, 140, 70)
   
   ' Create a path gradient based on the ellipse
   Call GdipCreatePathGradientFromPath(path, brush)
   
   ' Set the middle color of the path to Blue
   Call GdipSetPathGradientCenterColor(brush, Blue)
   
   ' Set the entire path boundary to Aqua
   ' NOTE: This expects an array, but since we only have one item we can fudge it
   Call GdipSetPathGradientSurroundColorsWithCount(brush, Aqua, 1)
   
   ' Draw the ellipse, keeping the exact coords we defined for the path
   Call GdipFillEllipse(graphics, brush, 0, 0, 140, 70)
   
   'Cleanup
   Call GdipDeletePath(path)     ' Delete the path object
   Call GdipDeleteBrush(brush)
   Call GdipDeleteGraphics(graphics)
End Sub

' NOTE: Use this same concept for: BMP, GIF, and PNG format saving
Private Sub BMPtoPNG()
   Dim img As Long, encoderCLSID As CLSID
   Dim stat As GpStatus

   ' Initializations
   ' No graphics object needed here since we aren't doing any drawing.
   ' We'll convert the grapes bitmap file
   Call GdipLoadImageFromFile(StrConv(App.path & "\GrapeBunch.bmp", vbUnicode), img)

   ' Get the CLSID of the PNG encoder
   Call GetEncoderClsid("image/png", encoderCLSID)

   ' Save as a PNG file. There are no encoder parameters for PNG images, so we pass a NULL.
   ' NOTE: The NULL (aka 0) must be passed byval, as the function declaration would get a pointer to the number 0.
   stat = GdipSaveImageToFile(img, StrConv(App.path & "\GrapeBunch.png", vbUnicode), encoderCLSID, ByVal 0)
   
   ' See if it was created
   If stat = Ok Then
      MsgBox "Successfully saved GrapeBunch.png!", vbInformation
   Else
      MsgBox "Error saving file! Status Code: " & stat, vbCritical
   End If
   
   ' Cleanup
   Call GdipDisposeImage(img)
End Sub

' Note: Use this same concept for: JPEG and TIFF saving
'       Also, it seems that you can pass a NULL for the EncoderParameters if you loaded a JPEG/TIFF and saving
'       the image back to the same format. If this is done, the current image properties should stay intact,
'       but I did not test that extensively.
Private Sub BMPtoJPEG()
   Dim img As Long, encoderCLSID As CLSID
   Dim stat As GpStatus
   Dim encoderParams As EncoderParameters
   Dim lngQuality As Long

   ' Initializations
   ' No graphics object needed here since we aren't doing any drawing.
   ' We'll convert the grapes bitmap file
   Call GdipLoadImageFromFile(StrConv(App.path & "\GrapeBunch.bmp", vbUnicode), img)
   
   ' Get the CLSID of the JPEG encoder
   Call GetEncoderClsid("image/jpeg", encoderCLSID)

   ' Save as a JPEG file. This format requires encoder parameters.
   lngQuality = 90   ' Quality is 90% of original
   ' Setup the encoder paramters
   encoderParams.count = 1    ' Only one element in this Parameter array
   With encoderParams.Parameter
      .NumberOfValues = 1     ' Should be one
      .type = EncoderParameterValueTypeLong
      ' Set the GUID to EncoderQuality
      .GUID = DEFINE_GUID(EncoderQuality)
      .value = VarPtr(lngQuality)  ' Remember: The value expects only pointers!
   End With
   
   ' Now save the bitmap as a jpeg at 10% compression
   stat = GdipSaveImageToFile(img, StrConv(App.path & "\GrapeBunch.jpg", vbUnicode), encoderCLSID, encoderParams)

   ' See if it was created
   If stat = Ok Then
      MsgBox "Successfully saved GrapeBunch.jpg!", vbInformation
   Else
      MsgBox "Error saving file! Status Code: " & stat, vbCritical
   End If
   
   ' Cleanup
   Call GdipDisposeImage(img)
End Sub

' Now that we know how to set the value of one encoding parameter, what do we do if we
' want to set more than one encoding parameter? Well, this function will show you how to
' do it!
' Note: Requires the CopyMemory API
Private Sub BMPtoJPEG_Params()
   Dim img As Long, encoderCLSID As CLSID
   Dim stat As GpStatus
   Dim encoderParams As EncoderParameters ' This will now become a temporary holder
   Dim encoderArray() As Byte             ' Our main "struct"
   Dim lngEP As Long                      ' Size of encoderParams variable/struct
   Dim lngQuality As Long

   ' Initializations
   ' No graphics object needed here since we aren't doing any drawing.
   ' We'll rotate the GrapeBunch.jpg file
   Call GdipLoadImageFromFile(StrConv(App.path & "\GrapeBunch.jpg", vbUnicode), img)
   lngEP = Len(encoderParams)
   
   ' Get the CLSID of the JPEG encoder
   Call GetEncoderClsid("image/jpeg", encoderCLSID)

   ' Determine how many parameters we will need
   ' JPEGs can only use 2 parameters, so we'll use both
   ReDim encoderArray(0 To (lngEP + Len(encoderParams.Parameter))) As Byte

   ' Save as a JPEG file. This format requires encoder parameters.
   ' NOTE: The quality here may not be taken into account for some reason; maybe due to the lossless rotation.
   lngQuality = 100   ' Quality is 100% of original
   ' Setup the encoder paramters
   ' We'll setup the struct and first parameter as usual
   encoderParams.count = 2    ' We are setting 2 parameters
   With encoderParams.Parameter
      .NumberOfValues = 1     ' Should be one
      .type = EncoderParameterValueTypeLong
      ' Set the GUID to EncoderQuality
      .GUID = DEFINE_GUID(EncoderQuality)
      .value = VarPtr(lngQuality)  ' Remember: The value expects only pointers!
   End With

   ' Copy the data into the byte array
   CopyMemory encoderArray(0), encoderParams, lngEP
   
   ' Now we'll re-use the parameter member of encoderParams
   With encoderParams.Parameter
      .NumberOfValues = 1     ' Should be one
      .type = EncoderParameterValueTypeLong
      ' Set the GUID to EncoderTransformation
      .GUID = DEFINE_GUID(EncoderTransformation)
      ' We'll flip horizontally - REMEMBER TO USE A POINTER!
      .value = VarPtr(EncoderValueTransformRotate180)
   End With

   ' Copy the second parameter to the byte array at the right offset
   CopyMemory encoderArray(lngEP), encoderParams.Parameter, Len(encoderParams.Parameter)
   
   ' Now save the bitmap as a jpeg at 0% compression to try to keep the quality up
   ' Notice how the byte array is passed instead of the struct
   stat = GdipSaveImageToFile(img, StrConv(App.path & "\GrapeBunch180.jpg", vbUnicode), encoderCLSID, encoderArray(0))

   ' See if it was created
   If stat = Ok Then
      MsgBox "Successfully saved GrapeBunch.jpg!", vbInformation
   Else
      MsgBox "Error saving file! Status Code: " & stat, vbCritical
   End If
   
   ' Cleanup
   Erase encoderArray
   Call GdipDisposeImage(img)
End Sub

Private Sub DrawCachedBitmap()
   Dim graphics As Long, bitmap As Long, cBitmap As Long
   Dim lngHeight As Long, lngWidth As Long
   Dim J As Long, k As Long

   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipLoadImageFromFile(StrConv(App.path & "\Texture.bmp", vbUnicode), bitmap)  ' Load the image
   ' Create a cached bitmap from the loaded image
   Call GdipCreateCachedBitmap(bitmap, graphics, cBitmap)

   ' Get the image height and width
   Call GdipGetImageHeight(bitmap, lngHeight)
   Call GdipGetImageWidth(bitmap, lngWidth)
   
   ' Perform a test to see which is faster
   For J = 1 To 300 Step 10
      For k = 1 To 1000
         Call GdipDrawImageRect(graphics, bitmap, J, J / 2, lngWidth, lngHeight)
      Next
   Next
   
   For J = 1 To 300 Step 10
      For k = 1 To 1000
         Call GdipDrawCachedBitmap(graphics, cBitmap, J, 150 + J / 2)
      Next
   Next

   ' Cleanup
   Call GdipDisposeImage(bitmap)
   Call GdipDeleteCachedBitmap(cBitmap)   ' Note the special deletion function
   Call GdipDisposeImage(cBitmap)         ' This may not be needed...
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawAlphaLines()
   Dim graphics As Long, bitmap As Long
   Dim lngHeight As Long, lngWidth As Long
   Dim opaquePen As Long, semiTansPen As Long

   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipLoadImageFromFile(StrConv(App.path & "\Texture.bmp", vbUnicode), bitmap)  ' Load the image
   ' Create our pens for line drawing
   Call GdipCreatePen1(ColorARGB(255, 0, 0, 255), 15, UnitPixel, opaquePen)
   Call GdipCreatePen1(ColorARGB(128, 0, 0, 255), 15, UnitPixel, semiTansPen) ' Has 50% alpha blending

   ' Get the image height and width
   Call GdipGetImageHeight(bitmap, lngHeight)
   Call GdipGetImageWidth(bitmap, lngWidth)

   ' Draw the image without auto-scaling
   Call GdipDrawImageRect(graphics, bitmap, 10, 5, lngWidth, lngHeight)
   
   ' Draw an opaque line over the image
   Call GdipDrawLine(graphics, opaquePen, 0, 20, 100, 20)
   ' Draw the semi-transparent line over the image
   Call GdipDrawLine(graphics, semiTansPen, 0, 40, 100, 40)
   ' Draw the same semi-transparent line, but with gamma correction
   Call GdipSetCompositingQuality(graphics, CompositingQualityGammaCorrected)
   Call GdipDrawLine(graphics, semiTansPen, 0, 60, 100, 60)
   
   ' Cleanup
   Call GdipDeletePen(opaquePen)
   Call GdipDeletePen(semiTansPen)
   Call GdipDisposeImage(bitmap)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawColorMatrix()
   Dim graphics As Long, bitmap As Long, pen As Long
   Dim imgAttr As Long, clrMatrix As ColorMatrix
   Dim lngHeight As Long, lngWidth As Long
   
   
   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipLoadImageFromFile(StrConv(App.path & "\Texture.bmp", vbUnicode), bitmap)   ' Load the image
   Call GdipCreatePen1(Black, 15, UnitPixel, pen)  ' Create an opaque pen

   ' Get the image height and width
   Call GdipGetImageHeight(bitmap, lngHeight)
   Call GdipGetImageWidth(bitmap, lngWidth)

   ' Fill the color matrix
   ' Notice the value 0.8 in row 4, column 4.
   clrMatrix.m(0, 0) = 1: clrMatrix.m(1, 0) = 0: clrMatrix.m(2, 0) = 0: clrMatrix.m(3, 0) = 0: clrMatrix.m(4, 0) = 0
   clrMatrix.m(0, 1) = 0: clrMatrix.m(1, 1) = 1: clrMatrix.m(2, 1) = 0: clrMatrix.m(3, 1) = 0: clrMatrix.m(4, 1) = 0
   clrMatrix.m(0, 2) = 0: clrMatrix.m(1, 2) = 0: clrMatrix.m(2, 2) = 1: clrMatrix.m(3, 2) = 0: clrMatrix.m(4, 2) = 0
   clrMatrix.m(0, 3) = 0: clrMatrix.m(1, 3) = 0: clrMatrix.m(2, 3) = 0: clrMatrix.m(3, 3) = 0.8: clrMatrix.m(4, 3) = 0
   clrMatrix.m(0, 4) = 0: clrMatrix.m(1, 4) = 0: clrMatrix.m(2, 4) = 0: clrMatrix.m(3, 4) = 0: clrMatrix.m(4, 4) = 1

   ' Create the ImageAttributes object
   Call GdipCreateImageAttributes(imgAttr)
   ' And set its color matrix
   Call GdipSetImageAttributesColorMatrix(imgAttr, ColorAdjustTypeDefault, True, clrMatrix, ByVal 0, ColorMatrixFlagsDefault)

   ' Draw a wide black line
   Call GdipDrawLine(graphics, pen, 10, 35, 200, 35)
   
   ' Draw the semi-transparent image
   Call GdipDrawImageRectRectI(graphics, bitmap, 30, 0, lngWidth, lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel, imgAttr)

   ' Cleanup
   Call GdipDisposeImageAttributes(imgAttr)  ' Delete the Image attributes object
   Call GdipDeletePen(pen)
   Call GdipDisposeImage(bitmap)
   Call GdipDeleteGraphics(graphics)
End Sub

' The slower way of using alpha-blending
Private Sub DrawAlphaPixels()
   Dim graphics As Long, bitmap As Long, pen As Long
   Dim lngHeight As Long, lngWidth As Long
   Dim iRow As Long, iCol As Long, lARGB As Long
   
   
   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipLoadImageFromFile(StrConv(App.path & "\Texture.bmp", vbUnicode), bitmap)  ' Load the image
   Call GdipCreatePen1(Black, 15, UnitPixel, pen)  ' Create an opaque pen

   ' Get the image height and width
   Call GdipGetImageHeight(bitmap, lngHeight)
   Call GdipGetImageWidth(bitmap, lngWidth)

   ' Modify the pixels in the bitmap
   ' NOTE: I'm pretty sure that the bitmap object it forever modified by doing this.
   '       If you still want the original, I would suggest cloning this image first.
   For iRow = 0 To (lngHeight - 1)
      For iCol = 0 To (lngWidth - 1)
         ' Get the current ARGB color of the pixel
         Call GdipBitmapGetPixel(bitmap, iCol, iRow, lARGB)
         ' Set the pixel color back with a new alpha
         ' NOTE: I created a helper function for alpha setting to make it easier
         Call GdipBitmapSetPixel(bitmap, iCol, iRow, ColorSetAlpha(lARGB, 255 * iCol / lngWidth))
      Next
   Next

   ' Draw a wide black line
   Call GdipDrawLine(graphics, pen, 10, 35, 200, 35)
   
   ' Draw the modified image
   Call GdipDrawImageRect(graphics, bitmap, 30, 0, lngWidth, lngHeight)
   

   ' Cleanup
   Call GdipDeletePen(pen)
   Call GdipDisposeImage(bitmap)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawAntiAliasText()
   Dim graphics As Long, brush As Long
   Dim fontFam As Long, curFont As Long
   Dim rcLayout As RECTF   ' Designates the string drawing bounds
   
   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipCreateSolidFill(Blue, brush)    ' Create a brush to draw the text with
   ' Create a font family object to allow use to create a font
   ' We have no font collection here, so pass a NULL for that parameter
   Call GdipCreateFontFamilyFromName(StrConv("Times New Roman", vbUnicode), 0, fontFam)
   ' Create the font from the specified font family name
   Call GdipCreateFont(fontFam, 32, FontStyleRegular, UnitPixel, curFont)
   
   ' Set up a drawing area
   ' NOTE: Leaving the right and bottom values at zero means there is no boundary
   rcLayout.Left = 10
   rcLayout.Top = 10
   
   ' This function allows us to alter the text quality.
   ' We'll use the worst quality first.
   Call GdipSetTextRenderingHint(graphics, TextRenderingHintSingleBitPerPixel)
   ' We have no string format object, so pass a NULL for that parameter
   Call GdipDrawString(graphics, StrConv("SingleBitPerPixel", vbUnicode), 17, curFont, rcLayout, 0, brush)
   
   
   ' Set up another drawing area
   rcLayout.Left = 10
   rcLayout.Top = 60
   
   ' Now we'll use anti-aliasing
   Call GdipSetTextRenderingHint(graphics, TextRenderingHintAntiAlias)
   ' We have no string format object, so pass a NULL for that parameter
   Call GdipDrawString(graphics, StrConv("AntiAlias", vbUnicode), 9, curFont, rcLayout, 0, brush)
   
   
   ' Cleanup
   Call GdipDeleteFont(curFont)     ' Delete the font object
   Call GdipDeleteFontFamily(fontFam)  ' Delete the font family object
   Call GdipDeleteBrush(brush)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawFormatText()
   Dim graphics As Long, brush As Long, pen As Long
   Dim fontFam As Long, curFont As Long, strFormat As Long
   Dim rcLayout As RECTF   ' Designates the string drawing bounds
   Dim str As String
   
   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipCreateSolidFill(Blue, brush)    ' Create a brush to draw the text with
   ' Create a font family object to allow use to create a font
   ' We have no font collection here, so pass a NULL for that parameter
   Call GdipCreateFontFamilyFromName(StrConv("Arial", vbUnicode), 0, fontFam)
   ' Create the font from the specified font family name
   ' >> Note that we have changed the drawing Unit from pixels to points!!
   Call GdipCreateFont(fontFam, 12, FontStyleUnderline + FontStyleBoldItalic, UnitPoint, curFont)
   ' Create the StringFormat object
   ' We can pass NULL for the flags and language id if we want
   Call GdipCreateStringFormat(0, 0, strFormat)
   
   ' Set up the drawing area boundary
   rcLayout.Left = 30
   rcLayout.Top = 10
   rcLayout.Right = 120
   rcLayout.Bottom = 140
   
   ' Center-justify each line of text
   Call GdipSetStringFormatAlign(strFormat, StringAlignmentCenter)
   
   ' Center the block of text (top to bottom) in the rectangle.
   Call GdipSetStringFormatLineAlign(strFormat, StringAlignmentCenter)
   
   ' Draw the string within the boundary
   str = StrConv("Use StringFormat and RectF objects to center text in a rectangle.", vbUnicode)
   Call GdipDrawString(graphics, str, -1, curFont, rcLayout, strFormat, brush)
   
   ' Create a pen and draw the boundary around the text
   Call GdipCreatePen1(Black, 1, UnitPixel, pen)
   Call GdipDrawRectangles(graphics, pen, rcLayout, 1)
   
   ' Cleanup
   Call GdipDeletePen(pen)
   Call GdipDeleteStringFormat(strFormat)
   Call GdipDeleteFont(curFont)     ' Delete the font object
   Call GdipDeleteFontFamily(fontFam)  ' Delete the font family object
   Call GdipDeleteBrush(brush)
   Call GdipDeleteGraphics(graphics)
End Sub

' Note: This example was inspired by another post on planetsourcecode today.
'       Someone asked if GDI+ could rotate images, and behold!
'       There are also several other ways to rotate an image.
Private Sub DrawRotated()
   Dim graphics As Long, img As Long, pen As Long
   Dim lngHeight As Long, lngWidth As Long
   
   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipLoadImageFromFile(StrConv(App.path & "\GrapeBunch.bmp", vbUnicode), img)

   ' Get the image height and width
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)

   ' This will rotate EVERYTHING!
   ' There are several rotation APIs available for you!
   Call GdipRotateWorldTransform(graphics, 45, MatrixOrderAppend)
   
   ' Make sure to provide a good x,y starting point!
   Call GdipDrawImageRect(graphics, img, 200, -150, lngWidth, lngHeight)

   ' Cleanup
   Call GdipDisposeImage(img)
   Call GdipDeleteGraphics(graphics)
End Sub


' NOTE: So many vars to use the text drawing!
'       A class for string encapsulation might be a wiser idea, depending on your needs.
Private Sub DrawClippedText()
   Dim graphics As Long, brush As Long, pen As Long
   Dim path As Long, polyPoints(1 To 4) As POINTL
   Dim region As Long, str As String
   Dim rcLayout As RECTF   ' Designates the string drawing bounds
   Dim fontFam As Long, curFont As Long


   ' Initialization
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   str = "A clipping region."  ' The demo text


   ' Create a path that consists of a single polygon
   ' Set the polygon points
   polyPoints(1).x = 10
   polyPoints(1).y = 10
   polyPoints(2).x = 150
   polyPoints(2).y = 10
   polyPoints(3).x = 100
   polyPoints(3).y = 75
   polyPoints(4).x = 100
   polyPoints(4).y = 150
   

   ' Create the path object and add the polygon to it
   Call GdipCreatePath(FillModeAlternate, path)
   Call GdipAddPathPolygonI(path, polyPoints(1), 4)
   
   ' Now create a region object based on the path
   ' The region object will allow us to set the clipping area/region
   Call GdipCreateRegionPath(path, region)
   
   ' Set the clipping region
   ' The default combine mode is CombineModeIntersect
   Call GdipSetClipRegion(graphics, region, CombineModeIntersect)


   ' Create a pen to draw the clipping region outline
   ' NOTE: The border looks a bit odd with 1 pixel width
   Call GdipCreatePen1(Black, 2, UnitPixel, pen)
   ' Draw the outline based on the path
   ' NOTE: You could also use GdipDrawPolygon if you wanted
   Call GdipDrawPath(graphics, pen, path)
   
   
   
   ' Create a font family object to allow use to create a font
   ' We have no font collection here, so pass a NULL for that parameter
   Call GdipCreateFontFamilyFromName(StrConv("Arial", vbUnicode), 0, fontFam)
   ' Create the font from the specified font family name
   Call GdipCreateFont(fontFam, 36, FontStyleBold, UnitPixel, curFont)
   ' Create a solid brush to draw the text with
   Call GdipCreateSolidFill(Red, brush)
   
   ' Draw the text twice inside the clipping region at different points
   rcLayout.Left = 15
   rcLayout.Top = 25
   Call GdipDrawString(graphics, StrConv(str, vbUnicode), Len(str), curFont, rcLayout, 0, brush)
   
   rcLayout.Left = 15
   rcLayout.Top = 68
   Call GdipDrawString(graphics, StrConv(str, vbUnicode), Len(str), curFont, rcLayout, 0, brush)

   
   ' Cleanup
   Call GdipDeleteBrush(brush)
   Call GdipDeletePen(pen)
   Call GdipDeletePath(path)
   Call GdipDeleteRegion(region)
   Call GdipDeleteFontFamily(fontFam)
   Call GdipDeleteFont(curFont)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawClippedImage()
   Dim graphics As Long, pen As Long, img As Long
   Dim path As Long, polyPoints(1 To 4) As POINTL
   Dim region As Long


   ' Initialization
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   ' Load the trusty GrapeBunch.bmp file - but using the Picture Box again!
   ' Load the bitmap file into the Picture box (could also embed it)
   Set picGrapes.Picture = LoadPicture("GrapeBunch.bmp")
   ' Get the image "class" from the PictureBox
   ' NOTE: This method uses the Picture object to get the needed handles.
   Call GdipCreateBitmapFromHBITMAP(picGrapes.Picture.Handle, picGrapes.Picture.hpal, img)
   ' Below is the "cheap" way (via file); good for all supported file type
   ' Could also use GdipCreateBitmapFromFile for this bitmap
   ' Comment out the picture box code above and uncomment this to try it out if you want!
   'Call GdipLoadImageFromFile(StrConv(app.path &"\GrapeBunch.bmp", vbUnicode), img)


   ' Create a path that consists of a single polygon
   ' Set the polygon points
   polyPoints(1).x = 10
   polyPoints(1).y = 10
   polyPoints(2).x = 150
   polyPoints(2).y = 10
   polyPoints(3).x = 100
   polyPoints(3).y = 75
   polyPoints(4).x = 100
   polyPoints(4).y = 150
   

   ' Create the path object and add the polygon to it
   Call GdipCreatePath(FillModeAlternate, path)
   Call GdipAddPathPolygonI(path, polyPoints(1), 4)
   
   ' Now create a region object based on the path
   ' The region object will allow us to set the clipping area/region
   Call GdipCreateRegionPath(path, region)
   
   ' Set the clipping region
   ' The default combine mode is CombineModeIntersect
   Call GdipSetClipRegion(graphics, region, CombineModeIntersect)


   ' Create a pen to draw the clipping region outline
   ' NOTE: The border looks a bit odd with 1 pixel width
   Call GdipCreatePen1(Black, 2, UnitPixel, pen)
   ' Draw the outline based on the path
   ' NOTE: You could also use GdipDrawPolygon if you wanted
   Call GdipDrawPath(graphics, pen, path)
   
   ' This will draw the image with auto-scaling, but since we won't be able to
   '  see the entire image, it won't matter here. The extra size will ensure that
   '  the entire clipping area will be visible.
   Call GdipDrawImageI(graphics, img, 0, 0)
   
   ' Cleanup
   Call GdipDisposeImage(img)
   Call GdipDeletePen(pen)
   Call GdipDeletePath(path)
   Call GdipDeleteRegion(region)
   Call GdipDeleteGraphics(graphics)
End Sub


' NOTE: This is a VERY rough demo of the HitTesting.
'       You should be able to make a much better version for your projects.
Private Sub DrawHitTest(Optional ByVal x As Long = 0, Optional ByVal y As Long = 0)
   Dim graphics As Long, brush As Long
   Dim region1 As Long, region2 As Long
   Dim rcRgn As RECTL
   Dim lngVisible As Long
   Static blnNotHit As Boolean   ' Static var determining if the last point was a hit. Saves a redraw.


   ' Set this to form var to true so that form's mouse move event will call this function.
   bHitTesting = True

   ' Initialization
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing

   ' Create a black brush - this is to go strictly by the SDK sample.
   ' We could just as easy create the brush later on; you'll see what I mean.
   Call GdipCreateSolidFill(Black, brush)
   
   ' Create the first region
   rcRgn.Left = 50
   rcRgn.Top = 0
   rcRgn.Right = 50
   rcRgn.Bottom = 150
   Call GdipCreateRegionRectI(rcRgn, region1)

   ' We can reuse rcRgn to create the second region
   rcRgn.Left = 0
   rcRgn.Top = 50
   rcRgn.Right = 150
   rcRgn.Bottom = 50
   Call GdipCreateRegionRectI(rcRgn, region2)
   
   ' Create a plus-shaped region by the union of region1 and region2.
   ' The union will replace region1.
   Call GdipCombineRegionRegion(region1, region2, CombineModeUnion)
   
   
   ' Assume that the "point" contains the location of the most recent click.
   ' To simulate a hit, assign (60, 10) to the point.
   ' To simulate a miss, assign (0, 0) to the point.
   Call GdipIsVisibleRegionPoint(region1, x, y, graphics, lngVisible)
   If lngVisible Then
      ' The point is in the region. Use an opaque brush.
      Call GdipSetSolidFillColor(brush, Red)
      
      ' Draw the region with the brush, as needed
      If blnNotHit = True Then
         Cls
         Call GdipFillRegion(graphics, brush, region1)
      End If
   Else
      ' The point is not in the region. Use a semitransparent brush.
      Call GdipSetSolidFillColor(brush, ColorARGB(68, 255, 0, 0))
      
      ' Draw the region with the brush, as needed
      If blnNotHit = False Then
         Cls
         Call GdipFillRegion(graphics, brush, region1)
      End If
   End If
   
   ' Update the local static var
   blnNotHit = Not CBool(lngVisible)
   
   ' Cleanup
   Call GdipDeleteBrush(brush)
   Call GdipDeleteRegion(region1)
   Call GdipDeleteRegion(region2)
   Call GdipDeleteGraphics(graphics)
End Sub

Private Sub DrawBezier()
   Dim graphics As Long, pen As Long
   
   ' Initialization
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   ' Create a nice colorful pen to draw with
   Call GdipCreatePen1(MediumAquamarine, 2, UnitPixel, pen)

   ' Add some anti-alias!
   Call GdipSetSmoothingMode(graphics, SmoothingModeAntiAlias)
   ' Draw the bezier line
   Call GdipDrawBezierI(graphics, pen, 10, 100, 100, 10, 150, 150, 200, 100)

   ' Cleanup
   Call GdipDeletePen(pen)
   Call GdipDeleteGraphics(graphics)
End Sub

' Reused some code from DrawScaling
' This will hopefully show you how to alter images and then save them!
Private Sub SaveScaling()
   Dim graphics As Long, img As Long, encoderCLSID As CLSID
   Dim lngHeight As Long, lngWidth As Long
   Dim new_img As Long

   ' Initialization
   Call GdipLoadImageFromFile(StrConv(App.path & "\GrapeBunch.bmp", vbUnicode), img)
   ' We need an *easy* way to copy all the bitmap palette, resolution, etc. data over to the new
   '  image we are going to create. The easy API way is to create a graphics object from the image.
   '  Once we do that, we can do whatever we like with it, but we won't touch the original image in this case.
   Call GdipGetImageGraphicsContext(img, graphics) ' Create a graphics object from the bitmap - we can draw to this now if we want, but we don't
   Call GetEncoderClsid("image/bmp", encoderCLSID) ' We'll save the image as a BMP file; get the image encoder CLSID
   
   ' Get the image height and width
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   
   ' Now, we will create a compatible graphics object from the current image
   Call GdipCreateBitmapFromGraphics(0.6 * lngWidth, 0.6 * lngHeight, graphics, new_img)
   ' We are going to re-use our graphics object that we created from the original image.
   ' We are doing this to save memory, and we only needed that graphics object to create a bitmap that would have
   '   the same color, bit-depth, etc. as the original.
   ' NOTE: You could also use another graphics object variable if you wanted.
   Call GdipDeleteGraphics(graphics)   ' Delete the graphics object
   ' Now create a brand new graphics object so we can draw on the new, blank bitmap in memory!
   Call GdipGetImageGraphicsContext(new_img, graphics)


   ' NOTE: We are using the graphics object for the NEW image in memory.
   ' Shrink the image using high-quality interpolation.
   Call GdipSetInterpolationMode(graphics, InterpolationModeHighQualityBicubic)
   ' NOTE: We have set the position to 0, 0.
   '       We are also using the original image object as the source.
   Call GdipDrawImageRectRectI(graphics, img, 0, 0, 0.6 * lngWidth, 0.6 * lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel)

   ' NOTES: Since we are shrinking the entire image, we could just as well have called
   '        the GdipDrawImageRectI function, which would simplify things - but our goal must
   '        be to make life hellish!
   
   ' Now save the file
   ' Remember: the BMP format has no encoder parameters!
   Call GdipSaveImageToFile(new_img, StrConv(App.path & "\GrapeBunchSmall.bmp", vbUnicode), encoderCLSID, ByVal 0)

   ' Cleanup
   Call GdipDisposeImage(img)
   Call GdipDisposeImage(new_img)
   Call GdipDeleteGraphics(graphics)
End Sub

' TODO: Load your favorite typelib with IStorage and IStream interfaces and uncomment!
'       Remember to uncomment the GDI+ IStream functions also!
' Another ported SDK sample that was nestled deep within the docs, with some changes.
'Private Sub SaveStream()
'   Dim img As Long, encoderCLSID As CLSID
'   Dim memfile() As Byte, F As Integer
'   Dim storage As IStorage
'   Dim stream As IStream
'   Dim pst As STATSTG
'
'
'   Call GdipLoadImageFromFile(StrConv(app.path & "\GrapeBunch.bmp", vbUnicode), img)
'   Call GetEncoderClsid("image/png", encoderCLSID) ' We'll save the image as a PNG file; get the image encoder CLSID
'
'   ' Create an in-memory storage object for the soon-to-be-created stream
'   If StgCreateDocfile(StrConv(app.path & "\GrapeBunch", vbUnicode), STGM_CREATE Or STGM_SHARE_EXCLUSIVE Or STGM_READWRITE Or STGM_DELETEONRELEASE, 0, storage) <> 0 Then
'      MsgBox "Error creating Docfile.", vbCritical
'   ' Create the stream
'   ElseIf storage.CreateStream(StrConv(app.path & "\GrapeBunch.png", vbUnicode), STGM_READWRITE Or STGM_SHARE_EXCLUSIVE, 0, 0, stream) <> 0 Then
'      MsgBox "Error creating stream.", vbCritical
'   ' Now save the loaded image to memory using the specified encoding
'   ElseIf GdipSaveImageToStream(img, stream, encoderCLSID, ByVal 0) <> Ok Then
'      MsgBox "Error saving to stream.", vbCritical
'   Else ' Dump the stream to a file; you'll also use something like this to send via internet
'      ' Get the image length, in bytes
'      If stream.stat(pst, STATFLAG_NONAME) = 0 Then
'         ' We have the info we need
'         ' Copy the data into an array/string
'         ' If you have an image more than 2 GB in size, God help you...
'         ReDim memfile(pst.cbSize.dwLowDWord)
'         ' Set the reading at the beginning of the stream
'         Call stream.Seek(0, STREAM_SEEK_SET, 0)
'         ' Read every single byte
'         If stream.Read(memfile(0), pst.cbSize.dwLowDWord, 0) = 0 Then
'         ' If the read was successful, save to disk
'            On Error Resume Next ' In case the kill failed
'            ' If we don't remove the old file, things could get wild.
'            ' Binary mode tends to not remove unwritten portions of the file if the size will have decreased,
'            '   thus bloating it and could cause problems.
'            Kill "GrapeBunchStream.png"
'            F = FreeFile
'            Open "GrapeBunchStream.png" For Binary Access Write As #F
'            Put #F, , memfile ' Put knows how to handle the byte array
'            Close #F
'         End If
'      End If
'   End If
'
'   ' Cleanup
'   ' NOTE: I'm not sure if the typelib will clean these up, so I will do it manually just in case.
'   '       If an error occurs when releasing, don't release, and vice versa.
'   Call stream.Release
'   Call storage.Release
'   Call GdipDisposeImage(img)
'End Sub



' NOTE: This example uses the JPEG produced by the BMPtoJPEG example.
'       If you haven't or don't want to run that example, convert the BMP to a JPEG somehow,
'       or enter a new image path.
Private Sub GetAllProperties()
   Dim img As Long
   Dim totalBufferSize As Long, numProperties As Long
   Dim allItems() As PropertyItem
   Dim I As Long

   ' Load an image with image properties, such as a JPEG or TIFF
   Call GdipLoadImageFromFile(StrConv(App.path & "\GrapeBunch.jpg", vbUnicode), img)
   
   ' Find out how many property items are in the image, and find out the
   ' required size of the buffer that will receive those property items.
   If GdipGetPropertySize(img, totalBufferSize, numProperties) <> Ok Then
      MsgBox "Image file not found or could not get property buffer size. Cannot continue!"
   ElseIf totalBufferSize = 0 Then ' Ensure there is a property buffer
      MsgBox "Invalid image property buffer size!", vbCritical
   Else ' We have a decent buffer size
      ' Allocate the number of structures we need
      ' NOTE: The totalBufferSize does not usually equal (Len(PropertyItem) * numProperties)!
      ReDim allItems(1) ' We need this to calculate the size of each PropertyItem.
      ' Determine how many items we need, and add one extra since the value is rounded down.
      ' NOTE: Could also use Len(), but I was getting a higher number of structures which were not used for some reason...
      I = (totalBufferSize / LenB(allItems(1))) + 1
      ReDim allItems(1 To I)
   
      ' Save the data to the byte array
      Call GdipGetAllPropertyItems(img, totalBufferSize, numProperties, allItems(1))

      ' Debug.print all properties
      For I = 1 To numProperties
         Debug.Print "Property #" & I & " ID: " & Hex(allItems(I).propId)
      Next
   End If

   ' Cleanup
   Erase allItems
   Call GdipDisposeImage(img)
End Sub


' NOTE: This example uses the JPEG produced by the BMPtoJPEG example.
'       If you haven't or don't want to run that example, convert the BMP to a JPEG somehow,
'       or enter a new image path.
' ALSO: This is somwhat similar to the above function.
'       I also opted to use the over-sized property array instead of a byte buffer, though
'       I admit it makes the code a bit harder to understand (though all the code probably seems that way!)
Private Sub ChangeTitleProperty()
   Dim img As Long, encoderCLSID As CLSID
   Dim totalBufferSize As Long
   Dim item() As PropertyItem
   Dim I As Long, stat As GpStatus
   Dim strTitle() As Byte


   ' The default title we will write
   ' NOTE: Prepare for another "hack"! It begins...(sigh)...
   ' NOTE: This is a byte array! The byte array is used to hold the ANSI string, since there is no consistant
   '       method I found to get the ANSI pointer to a VB string. This works like a charm though!
   '       Also notice the Chr$(0). This is the *required* NULL terminator - don't leave the byte array without it!
   '       (Better to place it here than redim the array to add it later.)
   strTitle = StrConv("This is a JPEG?!" & Chr$(0), vbFromUnicode)
   ' Load an image with image properties, such as a JPEG or TIFF
   Call GdipLoadImageFromFile(StrConv(App.path & "\GrapeBunch.bmp", vbUnicode), img)
   Call GetEncoderClsid("image/jpeg", encoderCLSID) ' We'll save the image as a JPEG file; get the image encoder CLSID
   
   stat = GdipGetPropertyItemSize(img, PropertyTagImageTitle, totalBufferSize)
   ' See if the buffer size retrieval went well.
   ' NOTE: If there is no title, we will set one!
   '       The result should be PropertyNotFound if there is no existing property of the type we asked for.
   If stat <> Ok And stat <> PropertyNotFound Then
      MsgBox "Image file not found or could not get property buffer size. Cannot continue!"
   ElseIf stat = Ok Then   ' We found the title; retrieve and reset it!
      ' Allocate the number of structures we need
      ' NOTE: The totalBufferSize does not usually equal (Len(PropertyItem) * numProperties)!
      ReDim item(0) ' We need this to calculate the size of each PropertyItem.
      ' Determine how many items we need. Note that the value is rounded down.
      ' NOTE: Could also use Len(), but I was getting a higher number of structures which were not used for some reason...
      I = totalBufferSize / LenB(item(0))
      ' NOTE: Since 0 is the lowest array index in this function, we need not add an extra one to the
      '       result due to the way VB handles arrays.
      ReDim item(0 To I)
   
      ' NOTE: You should check the resulting status codes for errors.
      Call GdipGetPropertyItem(img, PropertyTagImageTitle, totalBufferSize, item(0))
   
      ' Display the original title for fun
      Debug.Print "Original Image Title: " & PtrToStrA(item(0).value)
   
      ' Fill in the property info
      With item(0)
         .propId = PropertyTagImageTitle
         .length = UBound(strTitle)
         .value = VarPtr(strTitle(0))
         .type = PropertyTagTypeASCII
      End With

      ' Print the new title for fun
      Debug.Print "New Image Title: " & PtrToStrA(item(0).value)

      ' Set 'n' Save
      Call GdipSetPropertyItem(img, item(0))
      Call GdipSaveImageToFile(img, StrConv(App.path & "\GrapeBunchTitle.jpg", vbUnicode), encoderCLSID, ByVal 0)

   Else ' We don't have a title, but we will set one now!
      ' Only allocate on item; we can only set one property at a time, and we don't need more here anyway.
      ReDim item(0)

      ' Fill in the property info
      With item(0)
         .propId = PropertyTagImageTitle
         .length = UBound(strTitle)
         .value = VarPtr(strTitle(0))
         .type = PropertyTagTypeASCII
      End With

      ' Print the new title for fun
      Debug.Print "Image Title: " & PtrToStrA(item(0).value)

      ' Set 'n' Save
      Call GdipSetPropertyItem(img, item(0))
      Call GdipSaveImageToFile(img, StrConv(App.path & "\GrapeBunchTitle.jpg", vbUnicode), encoderCLSID, ByVal 0)
   End If
   
   ' Cleanup
   Erase strTitle
   Erase item
   Call GdipDisposeImage(img)
End Sub

Private Sub RetrievePathData()
   Dim graphics As Long, pen As Long, path As Long, brush As Long
   Dim Points(1 To 5) As POINTL, I As Long
   
   ' Random values (From the C++ SDK Sample)
   Points(1).x = 200
   Points(1).y = 200
   Points(2).x = 250
   Points(2).y = 240
   Points(3).x = 200
   Points(3).y = 300
   Points(4).x = 300
   Points(4).y = 310
   Points(5).x = 250
   Points(5).y = 350
   
   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics)           ' Create a graphics object for drawing
   Call GdipCreatePen1(Aquamarine, 5, UnitPixel, pen) ' Create a pen to draw with
   Call GdipCreatePath(FillModeAlternate, path)       ' Create a path object

   ' Path Construction
   Call GdipAddPathLineI(path, 20, 100, 150, 200)
   Call GdipAddPathRectangle(path, 40, 30, 80, 60)
   Call GdipAddPathEllipse(path, 200, 30, 200, 100)
   Call GdipAddPathCurveI(path, Points(1), UBound(Points))
   Call GdipDrawPath(graphics, pen, path)             ' Draw the path


   ' Now we will want to get the path data.
   ' NOTE: We have to get the count value manually and declare two arrays based on the count value.
   '       Remember to use floating points (POINTF) for the point array or you'll lose precision!
   Dim pdata As PathData, bytebuf() As Byte, ptbuf() As POINTF
   
   ' I opted to reuse the pdata.count variable instead of allocating another variable to hopefully save a little memory.
   Call GdipGetPointCount(path, pdata.count)
   
   ' Resize the arrays.
   ReDim bytebuf(1 To pdata.count)
   ReDim ptbuf(1 To pdata.count)
   
   ' Assign the pointers to the PathData.
   pdata.Points = VarPtr(ptbuf(1))
   pdata.types = VarPtr(bytebuf(1))
   
   ' Now retrieve the points and types.
   Call GdipGetPathData(path, pdata)

   ' Draw the path's data points
   Call GdipCreateSolidFill(Red, brush)   ' Create a brush for the ellipses we are about to draw
   For I = 1 To pdata.count
      Call GdipFillEllipse(graphics, brush, ptbuf(I).x - 3, ptbuf(I).y - 3, 6#, 6#)
   Next

   ' Cleanup
   Erase bytebuf
   Erase ptbuf
   Call GdipDeleteBrush(brush)
   Call GdipDeletePen(pen)
   Call GdipDeletePath(path)
   Call GdipDeleteGraphics(graphics)
End Sub

' Play all the images of a GIF animation once.
Private Sub PlayGIF()
   Dim graphics As Long, img As Long, dID As CLSID
   Dim frameCount As Long, arDelay() As Long, delay As Long, loopCount As Long
   Dim lngHeight As Long, lngWidth As Long
   Dim I As Long, item() As PropertyItem, totalBufferSize As Long
   
   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics)           ' Create a graphics object for drawing
   Call GdipLoadImageFromFile(StrConv(App.path & "\Light.gif", vbUnicode), img)
   
   ' Get the image height and width
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   
   ' Load all of the needed information
   ' NOTE: Since GIF images are animations, they are stored in the Time dimension
   ' Get the GUID of the frame dimension we will be looking at
   dID = DEFINE_GUID(FrameDimensionTime)
   Call GdipImageGetFrameCount(img, dID, frameCount)

   ' Get the frame delay counts
   ' To get the delay times, we must delve into the property items.
   ' I'm going to over-dim the PropertyItem array instead of some fancy IMalloc or Byte array and CopyMemory.
   
   ' See if the buffer size retrieval went well.
   If GdipGetPropertyItemSize(img, PropertyTagFrameDelay, totalBufferSize) <> Ok Then
      MsgBox "Image file not found, could not get the property buffer size, or property not found. Cannot continue!"
   Else
      ' Allocate the number of structures we need
      ' NOTE: The totalBufferSize does not usually equal (Len(PropertyItem) * numProperties)!
      ReDim item(0) ' We need this to calculate the size of each PropertyItem.
      ' Determine how many items we need. Note that the value is rounded down.
      ' NOTE: Could also use Len(), but I was getting a higher number of structures which were not used for some reason...
      I = totalBufferSize / LenB(item(0))
      ' NOTE: Since 0 is the lowest array index in this function, we need not add an extra one to the
      '       result due to the way VB handles arrays.
      ReDim item(0 To I)
   
      ' NOTE: You should check the resulting status codes for errors.
      Call GdipGetPropertyItem(img, PropertyTagFrameDelay, totalBufferSize, item(0))
      
      ' Save the delay times
      ' The returned array will be one-based...ugh - no uniformity!
      arDelay = GetPropValue(item(0))

      ' Display the number of frames for fun
      Debug.Print "Frames: " & frameCount
      
      ' TODO: Use PropertyTagLoopCount to determine how many times to play the GIF
      
      ' Loop through the frames
      ' NOTE: We are assuming all frames are in this one dimension, as there can be several dimensions.
      ' ALSO: The frames are zero-based, while the count we retrieved is one-based.
      For I = 0 To frameCount - 1
         ' Select the current frame into the image object
         Call GdipImageSelectActiveFrame(img, dID, I)
         ' Now draw that frame
         Call GdipDrawImageRectI(graphics, img, 0, 0, lngWidth, lngHeight)
   
         ' Delay
         ' NOTE: You should use the Multimedia Timer Functions for production apps.
         ' ALSO: The delay is in hundredths (1/100) of a second.
         delay = GetTickCount
         Do While GetTickCount < delay + (arDelay(I + 1) * 10) ' Multiply by 10 to convert to milliseconds
            DoEvents ' This is probably not the best stalling technique
         Loop
      Next
   End If

On Error Resume Next ' Just in case the Erase poses raises an error when erasing an empty array
   ' Cleanup
   Erase arDelay
   Erase item
   Call GdipDisposeImage(img)
   Call GdipDeleteGraphics(graphics)
End Sub

' Crop an image into a new bitmap.
' NOTE: If you only want to draw a cropped image, just shorten the Width/Height
'       parameters when calling the DrawImage* function(s).
' ALSO: There is more than one way to create a new cropped image with GDI+!
Private Sub CropImage()
   Dim img As Long, imgCrop As Long
   Dim encoderCLSID As CLSID, stat As GpStatus
   Dim lngWidth As Long, lngPixelFormat As Long


   ' Load a file to crop.
   Call GdipLoadImageFromFile(StrConv(App.path & "\GrapeBunch.bmp", vbUnicode), img)
   ' Get the image encoder CLSID for saving.
   ' I'll save the cropped image as a PNG for simplicity.
   Call GetEncoderClsid("image/png", encoderCLSID)
   
   ' Get the image width; I don't need the height.
   Call GdipGetImageWidth(img, lngWidth)
   
   ' Get the current image pixel format.
   ' The C++ SDK clone example used PixelFormatDontCare, but this can limit what you can do with the image.
   Call GdipGetImagePixelFormat(img, lngPixelFormat)

   ' Create the cropped image.
   Call GdipCloneBitmapAreaI(0, 0, lngWidth, 205, lngPixelFormat, img, imgCrop)
   
   ' Do whatever you will with the cropped image!
   ' I'll only save to a file.
   stat = GdipSaveImageToFile(imgCrop, StrConv(App.path & "\GrapeCrop.png", vbUnicode), encoderCLSID, ByVal 0)
   If stat = Ok Then
      MsgBox "Cropped image successfully saved.", vbInformation
   Else
      MsgBox "Error saving cropped image!", vbCritical
   End If
   
   ' Cleanup
   Call GdipDisposeImage(img)
   Call GdipDisposeImage(imgCrop)
End Sub


' Grayscale matrix values courtesy: Dana Seaman
' NOTE: GDI+ v1.0 cannot save in true grayscale, but instead saves these grayscale
'       images in gray RGB colors.
Private Sub SaveGrayscale()
   Dim graphics As Long, img As Long, new_img As Long
   Dim lngHeight As Long, lngWidth As Long
   Dim imgAttr As Long, clrMatrix As ColorMatrix
   Dim encoderCLSID As CLSID, stat As GpStatus


   ' Initialization
   Call GdipLoadImageFromFile(StrConv(App.path & "\GrapeBunch.bmp", vbUnicode), img)
   Call GetEncoderClsid("image/bmp", encoderCLSID)
      
   ' Get the image height and width
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)
   
   ' All 9 matrix items must be the same value, so use of a constant or variable will ease things
   '  when you desire a contrast change.
   ' 1 = High contrast (very bright looking), 0 = Low contrast (totally black or very dark)
   ' NOTE: I hope "contrast" was the correct word I was looking for here...
   ' ALSO: You can still implement transparency via (3,3) if desired.
   Const sngContrast As Single = 0.35  ' 0.35 looks about the same as PSP7 making the image grayscale
   
   ' Fill the color matrix
   clrMatrix.m(0, 0) = sngContrast: clrMatrix.m(1, 0) = sngContrast: clrMatrix.m(2, 0) = sngContrast: clrMatrix.m(3, 0) = 0: clrMatrix.m(4, 0) = 0
   clrMatrix.m(0, 1) = sngContrast: clrMatrix.m(1, 1) = sngContrast: clrMatrix.m(2, 1) = sngContrast: clrMatrix.m(3, 1) = 0: clrMatrix.m(4, 1) = 0
   clrMatrix.m(0, 2) = sngContrast: clrMatrix.m(1, 2) = sngContrast: clrMatrix.m(2, 2) = sngContrast: clrMatrix.m(3, 2) = 0: clrMatrix.m(4, 2) = 0
   clrMatrix.m(0, 3) = 0: clrMatrix.m(1, 3) = 0: clrMatrix.m(2, 3) = 0: clrMatrix.m(3, 3) = 1: clrMatrix.m(4, 3) = 0
   clrMatrix.m(0, 4) = 0: clrMatrix.m(1, 4) = 0: clrMatrix.m(2, 4) = 0: clrMatrix.m(3, 4) = 0: clrMatrix.m(4, 4) = 1

   ' Create the ImageAttributes object to set the color matrix
   Call GdipCreateImageAttributes(imgAttr)
   Call GdipSetImageAttributesColorMatrix(imgAttr, ColorAdjustTypeDefault, True, clrMatrix, ByVal 0, ColorMatrixFlagsDefault)

   ' A slimy hack to fairly easily make a new image in memory with the same image properties
   ' Note that I am re-using the graphics variable in this function.
   Call GdipGetImageGraphicsContext(img, graphics)
   Call GdipCreateBitmapFromGraphics(lngWidth, lngHeight, graphics, new_img)
   Call GdipDeleteGraphics(graphics)   ' Cleanup so we can reuse the variable
   
   ' Now create a graphics object based on the new, blank memory bitmap
   Call GdipGetImageGraphicsContext(new_img, graphics)

   ' Draw the image in grayscale to the memory bitmap we will eventually save
   Call GdipDrawImageRectRectI(graphics, img, 0, 0, lngWidth, lngHeight, 0, 0, lngWidth, lngHeight, UnitPixel, imgAttr)

   ' Save the grayscale image
   ' Remember to pass that zero ByVal! (I kept forgetting and VB crashed many a time)
   stat = GdipSaveImageToFile(new_img, StrConv(App.path & "\GrapeBunch_Gray.bmp", vbUnicode), encoderCLSID, ByVal 0)
   If stat = Ok Then
      MsgBox "Successfully saved grayscale file!", vbInformation
   Else
      MsgBox "Error saving file! Status Code: " & stat, vbCritical
   End If
   

   ' Cleanup
   Call GdipDisposeImageAttributes(imgAttr)
   Call GdipDisposeImage(img)
   Call GdipDisposeImage(new_img)
   Call GdipDeleteGraphics(graphics)
End Sub

' Watch the mysterious vertical line drawing delay in action.
' The vertical lines are drawn more slowly than horizontal lines.
' If you can't see the difference 'as-is', comment out one of the for loops to see the actual speed difference of the other.
Private Sub DrawGrid()
   Dim graphics As Long, pen As Long
   Dim I As Long
   Const H As Long = 500
   Const W As Long = 500
   
   ' Init
   Call GdipCreateFromHDC(Me.hdc, graphics)
   Call GdipCreatePen1(Black, 1, UnitPixel, pen)

   For I = 0 To W Step 5
      Call GdipDrawLineI(graphics, pen, I, 0, I, H)
   Next
   
   For I = 0 To H Step 5
      Call GdipDrawLineI(graphics, pen, 0, I, W, I)
   Next

   ' Cleanup
   Call GdipDeletePen(pen)
   Call GdipDeleteGraphics(graphics)
End Sub

' Draw text with a specific background brush/color.
Private Sub DrawTextBackColor()
   Dim graphics As Long, brush As Long, pen As Long
   Dim fontFam As Long, curFont As Long, brushBG As Long
   Dim rcLayout As RECTF   ' Designates the string drawing bounds
   Dim rcOrigin As RECTF   ' Only the Left/X and Top/Y members are used; rest should be zero
   Dim str As String, strlen As Long
   
   ' Initializations
   str = "Text with a background color!"
   strlen = Len(str)
   str = StrConv(str, vbUnicode) ' Now convert to Unicode for GDI+
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipCreateSolidFill(Blue, brush)    ' Create a brush to draw the text with
   Call GdipCreateSolidFill(Red, brushBG)   ' Create a solid color brush to draw the text background with
   ' Create a font family object to allow use to create a font
   ' We have no font collection here, so pass a NULL for that parameter
   Call GdipCreateFontFamilyFromName(StrConv("Arial", vbUnicode), 0, fontFam)
   ' Create the font from the specified font family name
   ' >> Note that we have changed the drawing Unit from pixels to points!!
   Call GdipCreateFont(fontFam, 12, FontStyleUnderline + FontStyleBold, UnitPoint, curFont)

   ' (x,y) location to draw the string
   rcOrigin.Left = 60
   rcOrigin.Top = 20
   
   ' Get the size of the output text so we know how big of a background rectangle to draw
   ' NOTE: Could pass -1 for the string length, however, the returned rectangle will include
   '       space for the NULL character, which makes the rect larger than needed, hence strlen - 1
   '       to account for the NULL character.
   Call GdipMeasureString(graphics, str, strlen - 1, curFont, rcOrigin, 0, rcLayout, ByVal 0, ByVal 0)
   
   ' Draw the background color
   Call GdipFillRectangles(graphics, brushBG, rcLayout, 1)
   
   ' Draw the string
   Call GdipDrawString(graphics, str, -1, curFont, rcOrigin, 0, brush)
   
   ' Cleanup
   Call GdipDeleteFont(curFont)     ' Delete the font object
   Call GdipDeleteFontFamily(fontFam)  ' Delete the font family object
   Call GdipDeleteBrush(brush)
   Call GdipDeleteBrush(brushBG)
   Call GdipDeleteGraphics(graphics)
End Sub

' Write to an image using a User Input Buffer
Private Sub LockBitsWriteUIB()
   Dim graphics As Long
   Dim lngHeight As Long, lngWidth As Long
   Dim bitmap As Long
   Dim bmpData As BitmapData
   Dim rc As RECTL


   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics)
   ' Create a bitmap object from a BMP file
   Call GdipLoadImageFromFile(StrConv(App.path & "\LockBitsImage.bmp", vbUnicode), bitmap)
     
   ' Get the image height and width
   Call GdipGetImageHeight(bitmap, lngHeight)
   Call GdipGetImageWidth(bitmap, lngWidth)
   
   ' Create and fill a pixel data buffer.
   Dim pixels(30, 50) As Long
   Dim row As Long, col As Long
   ' Pixels we alter will all turn to Aqua
   While row < 30
      col = 0
      While col < 50
         pixels(row, col) = Aqua
         col = col + 1
      Wend
      row = row + 1
   Wend

   ' Since we are providing our own buffer, we need for fill in the BitmapData structure
   bmpData.Width = 50
   bmpData.Height = 30
   bmpData.stride = 4 * bmpData.Width
   bmpData.PixelFormat = PixelFormat32bppARGB
   bmpData.scan0 = VarPtr(pixels(0, 0))
   'bmpData.Reserved = 0 ' Not needed as VB zeros it out already

   ' Display the bitmap before alterations
   Call GdipDrawImageRect(graphics, bitmap, 10, 10, lngWidth, lngHeight)

   ' Constants specified by the GDI+ C++ SDK Example
   rc.Left = 20   ' Starting X coord within image
   rc.Top = 10    ' Starting Y coord within image
   rc.Right = 50  ' Width of locked pixel area
   rc.Bottom = 30 ' Height of locked pixel area

   ' Lock a 50×30 rectangular portion of the bitmap for writing.
   ' NOTE: We can't use a floating point rect here (no RECTF, only RECTL)...and yes, it matters.
   Call GdipBitmapLockBits(bitmap, rc, ImageLockModeWrite Or ImageLockModeUserInputBuf, PixelFormat32bppARGB, bmpData)
   ' INFO: Thank goodness for the ImageLockModeUserInputBuf flag!
   ' We would otherwise need to get a VB array to somehow point to the UINT (VB type Long) array of pixels.
   ' This would likely involve a VB array and CopyMemory, or something even more complicated if you desired
   ' to alter the pixels immediately without yet another CopyMemory call (I'm sure there is something
   ' someone made out there to make this sort of thing easier in VB though; I just don't know about it).
   
   ' Commit the changes and unlock the 50×30 portion of the bitmap.
   Call GdipBitmapUnlockBits(bitmap, bmpData)

   ' Display the altered bitmap.
   Call GdipDrawImageRect(graphics, bitmap, 150, 10, lngWidth, lngHeight)
   
   ' Cleanup
   Call GdipDisposeImage(bitmap)
   Call GdipDeleteGraphics(graphics)
End Sub


' Saves an image to a single-frame GIF because I'm feeling lazy.
' WARNING: This will save using the 256 color halftone palette!
'          See Q315780 for more info.
Private Sub BMPtoGIF()
   Dim img As Long, encoderCLSID As CLSID
   Dim stat As GpStatus


   ' Initializations
   ' No graphics object needed here since we aren't doing any drawing.
   ' We'll convert the grapes bitmap file
   Call GdipLoadImageFromFile(StrConv(App.path & "\GrapeBunch.bmp", vbUnicode), img)
   
   ' Get the CLSID of the GIF encoder
   Call GetEncoderClsid("image/gif", encoderCLSID)

   ' Now save the bitmap as a gif
   ' NOTE: The image will be saved using the Halftone palette, unless the image has its own palette,
   '       except under some circumstances. See Q318343 and Q315780 for more info.
   stat = GdipSaveImageToFile(img, StrConv(App.path & "\GrapeBunch.gif", vbUnicode), encoderCLSID, ByVal 0)

   ' See if it was created
   If stat = Ok Then
      MsgBox "Successfully saved GrapeBunch.gif!", vbInformation
   Else
      MsgBox "Error saving file! Status Code: " & stat, vbCritical
   End If
   
   ' Cleanup
   Call GdipDisposeImage(img)
End Sub


' Takes a paletted bitmap and makes the white color trasparent when saving the single-frame GIF.
' See Q318343 and Q315780 for more info.
Private Sub BMPtoGIF_Transparency()
   Dim img As Long, pfImg As Long, imgGIF As Long
   Dim encoderCLSID As CLSID
   Dim stat As GpStatus
   Dim palette As ColorPalette, palSize As Long
   Dim I As Long, rc As RECTL
   Dim bmpData As BitmapData
   Dim lngWidth As Long, lngHeight As Long


   ' Initializations
   ' Load the paletted image somehow
   Call GdipLoadImageFromFile(StrConv(App.path & "\GrapeBunchPaletted.bmp", vbUnicode), img)

   ' Get the CLSID of the GIF encoder
   Call GetEncoderClsid("image/gif", encoderCLSID)

   ' Get the image height and width
   Call GdipGetImageHeight(img, lngHeight)
   Call GdipGetImageWidth(img, lngWidth)

   ' Get the pixel format of the image
   Call GdipGetImagePixelFormat(img, pfImg)

   ' Check to see if the palette buffer is large enough
   Call GdipGetImagePaletteSize(img, palSize)
   If palSize <= LenB(palette) Then
      ' Create the new, blank image which will be the GIF.
      ' For some reason we need to create the new image to allow for the palette changes.
      ' The KB articles listed above may give you more insight.
      Call GdipCreateBitmapFromScan0(lngWidth, lngHeight, 0, pfImg, ByVal 0, imgGIF)
   
      ' Get the original image palette
      Call GdipGetImagePalette(img, palette, palSize)
      ' Ensure the palette recognizes the alpha
      palette.flags = palette.flags Or PaletteFlagsHasAlpha
      ' Ensure the color we want to be transparent is the only color with zero alpha
      For I = 0 To palette.count - 1 ' The count is one-based, but the array is zero-based; adjust.
         If palette.Entries(I) = Colors.White Then
            palette.Entries(I) = ColorSetAlpha(palette.Entries(I), 0)
         Else
            palette.Entries(I) = ColorSetAlpha(palette.Entries(I), 255)
         End If
      Next
      
      ' Set the palette of the soon-to-be GIF image
      Call GdipSetImagePalette(imgGIF, palette)
      
      ' Prepare to lock the bits on the original
      ' Fill the rect with the area we want (entire image)
      rc.Left = 0   ' Starting X coord within image
      rc.Top = 0    ' Starting Y coord within image
      rc.Right = lngWidth   ' Width of locked pixel area
      rc.Bottom = lngHeight ' Height of locked pixel area
      
      ' Reassign the pixels with as little effort possible since we don't want to access them, otherwise you might consider using
      '  a user input buffer.
      Call GdipBitmapLockBits(img, rc, ImageLockModeRead, pfImg, bmpData)
      Call GdipBitmapLockBits(imgGIF, rc, ImageLockModeWrite Or ImageLockModeUserInputBuf, pfImg, bmpData)
      ' WARNING: Close the locks in reverse order of lockage as shown here to prevent possible errors!
      Call GdipBitmapUnlockBits(imgGIF, bmpData)
      Call GdipBitmapUnlockBits(img, bmpData)
   
      ' Now save the bitmap as a gif
      stat = GdipSaveImageToFile(imgGIF, StrConv(App.path & "\GrapeBunch.gif", vbUnicode), encoderCLSID, ByVal 0)
   
      ' See if it was created
      If stat = Ok Then
         MsgBox "Successfully saved GrapeBunch.gif!", vbInformation
      Else
         MsgBox "Error saving file! Status Code: " & stat, vbCritical
      End If
   Else
      ' You would think 256 colors, each treated as a 4 byte Long would be large enough...but Noooo! Error!
      MsgBox "Palette buffer not large enough! Cannot process image!", vbCritical
   End If

   ' Cleanup
   Call GdipDisposeImage(img)
   Call GdipDisposeImage(imgGIF)
End Sub
