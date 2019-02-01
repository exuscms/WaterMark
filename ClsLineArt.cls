VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLineArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'--------------------------------------------------------'
' Line Art Creation, And Edge Dection By Arivnder Sehmi. '
' Arvinder@Sehmi.org.uk                                  '
' September 23th 2000                                    '
'Converted to Class by Roger Gilchrist December 10 2002  '
'rojagilkrist@hotmail.com                                '
'Added Grey-Scale code and supports routines Dec 23      '
'--------------------------------------------------------'
'FilterTo, Brighten, DarkenX,Lighten and Lighten1 were inspired by
'Steven Paterson at 'http://www.tolon.co.uk/ste/documents.html'
'which is his web-site where he goes into a lot of the maths
'NOTE Dec 2002 his green calculation is in error
'search for 'Graphical Routine Maths' at PSC for other comments on his ideas
'
Private sour As PictureBox 'Source PictureBox
Private Dest As PictureBox 'Destination PictureBox
Private DecimalPercentBox As Label ' Optional control - allow you to show DecimalPercentage of conversion on a label
'                             (You might change this to some other control type )
Private DecimalPercentBoxExists As Boolean 'Test that this optional control exists before attepmting to use it

'Internal Property Variables
Private m_Tolerance As Long 'hold the Tolerance value from form Def= 157
Private m_Invert As Boolean 'hold Invert Def= False
Private m_Red As Boolean    'Use Red in Greyscaling Def=True
Private m_Blue As Boolean   'Use Blue in Greyscaling Def=True
Private m_Green As Boolean  'Use Green in Greyscaling Def=True
Private m_ShowConversion As Boolean 'Allow continuous Updating of image Def=True
Private m_Working As Boolean 'Allows class to report that it is working
Private m_Cancel As Boolean ' has the cancel button been pressed?
Private m_Greys As Integer
Private m_Greylimit As Long
Private m_IgnoreTolerance As Boolean
Private m_GreyWeirdness As Boolean
Public Enum EdgeMode
    AllEdge
    LeftEdge
    RightEdge
    TopEdge
    BottomEdge
    LeftRightEdge
    TopBottomEdge
    LeftTopEdge
    LeftBottomEdge
    RightTopEdge
    RightBottomEdge
End Enum
Rem Mark off
'The comment above just stops Ulli's VB Code Formatter worrying about the following
'This is just a trick to force IDE to maintain the Case of the Enums
#If False Then
Dim AllEdge
Dim LeftEdge
Dim RightEdge
Dim TopEdge
Dim BottomEdge
Dim LeftRightEdge
Dim TopBottomEdge
Dim LeftTopEdge
Dim LeftBottomEdge
Dim RightTopEdge
Dim RightBottomEdge
#End If
Rem Mark on
Private m_GreyMode As Integer
Private m_PixelSet As Integer
Private m_AnyPixel As Boolean ''True = any Pixel is good enough 'False =Test all Pixels
Private m_ChangeWhileDrawing As Boolean
'Description
'SETUP
'Place following in Declaration section
'
'       Private|Public LA As New ClsLineArt
'
'Place following in Form_Load or Sub Main (whereever you set up initial values)
'
'       LA.AssignControls Source, Dest, DecimalPercentDone

'CREATE controls to set the Properties
'Boolean controls(OptionButton|CheckBox|Checked Menus) for Blue, Green, Invert, Red, ShowConversion
'Numerical for Tolerence (Range 0 - 255)
'AND to actually use it
'Buttons or Menu to start EdgeDetect and LineArt
'Button for Cancel

'ROUTINES (Type/Alphabetic)
'Public Subs
'AssignControls --  Assigns controls for class to use.By Roger Gilchrist.
'EdgeDetect     --  By Arivnder Sehmi. Does the EdgeDetection conversion. Modified by RG.
'GreyScaleImage --  Convert image to grey scale.
'LineArt        --  By Arivnder Sehmi. Does the LineArt conversion. Modified by RG.
'Experimental   --  Wrapper routine for testing other effects
'
'Public Properties  allow user to read/write various values for class (Note: some are Read or Write Only)
'AnyPixel
'Blue           --  used in LineArtTest to include/exclude Blue from grey conversion
'Cancel         --  Write Only. If set EdgeDetect and LineArt are aborted
'ChangeWhileDrawing  -- Write Only. Allows you to change Tolereance in EdgeDEtection while drawing
'Green          --  used in LineArtTest to include/exclude Green from grey conversion
'Invert         --  used in EdgeDetect and LineArt
'Red            --  used in LineArtTest to include/exclude Red from grey conversion
'ShowConversion --  'used in WorkReport(slight speed gain if False but not really worth setting)
'Tolerance      --  used in EdgeDetect and LineArt
'Working        --  Read Only. Allows other parts of program to test if EdgeDetect and LineArt are running
'
'Private Subs and Functions
'LineArtTest --  By Arivnder Sehmi. average colour to grey. Modified to return Boolean
'EdgeDetectTest --  By Arivnder Sehmi. Detects edges. Modified to return Boolean
'EdgeDetectTestStrong   --  By Roger Gilchrist. Calls EdgeDetectTest to allow different edges to be detected
'GreyPixel      --  convert long colour to long grey colour
'LongToRGB      --  support for GreyPixel
'Class_Initialize   sets the default values
'PreventAllOff  --  Makes sure that at least one of the colours is always available
'WorkDone       --  Encapsulate commeon parts of EdgeDetect and LineArt routines (Finish)
'WorkInitialize --  Encapsulate commeon parts of EdgeDetect and LineArt routines (Start)
'WorkReport     --  Encapsulate commeon parts of EdgeDetect and LineArt routines (Running)
'
''Experimental subs
'Smoother   --  Succeeded Uses PixelRing            averages out pixel value
'Sharper    --  Failed    Uses PixelRing            brightens pixel
'Sharper1   --  Succeeded Uses PixelRing            increases separation of pixel values
'NeonCrayon     --  Doubtful  Uses PixelRing & Tolerance   run to see
'Diffuse    --  Succeeded  Uses PixelRing & Tolerance  select a random neighbour pixel colour to use

'API calls that do the drawing. Arivnder Sehmi supplied this
Private Enum FilterTones
    fRed
    fGreen
    fBlue
    fYellow
    fPurple
    fAqua
End Enum
Rem Mark off
'The comment above just stops Ulli's VB Code Formatter worrying about the following
'This is just a trick to force IDE to maintain the Case of the Enums
#If False Then
Dim fRed
Dim fGreen
Dim fBlue
Dim fYellow
Dim fPurple
Dim fAqua
#End If
Rem Mark on

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Property Get AnyPixel() As Boolean

    AnyPixel = m_AnyPixel

End Property

Public Property Let AnyPixel(ByVal Value As Boolean)

  'True = any Pixel is good enough
  'False =Test all Pixels

    m_AnyPixel = Value

End Property

Public Sub AssignControls(P1 As PictureBox, P2 As PictureBox, Optional DecimalPercentLabel As Label)

  'set up controls to use

    Set sour = P1
    Set Dest = P2

    'You might like to change this to use a Progressbar
    If DecimalPercentLabel Is Nothing Then
        DecimalPercentBoxExists = False
      Else 'NOT DecimalPercentLABEL...
        DecimalPercentBoxExists = True
        Set DecimalPercentBox = DecimalPercentLabel
        DecimalPercentBox.Caption = ""
    End If

End Sub

Public Property Let Blue(ByVal BVal As Boolean)

  'used in LineArtTest to include/exclude Blue from grey conversion

    m_Blue = PreventAllOff(BVal)

End Property

Public Property Get Blue() As Boolean

    Blue = m_Blue

End Property

Private Function Brighten(OwnCol As Long, x As Single, Y As Single, TolFactor As Integer) As Long

  'based on Chapter 5 'Visual Basic Power Toolkit' by R. Mansfield and E. Peroutsos 1955

  Dim iRed As Integer, iGreen As Integer, iBlue As Integer
  Dim lngRed As Long, lngGreen As Long, lngBlue As Long
  Dim oRed As Integer, oGreen As Integer, oBlue As Integer
  Dim DivFactor As Integer
  Dim i As Integer
  Dim Tx(8) As Single
  Dim TY(8) As Single
  Dim L_Tol As Single

    LongToRGB OwnCol, iRed, iGreen, iBlue
    lngRed = iRed
    lngGreen = iGreen
    lngBlue = iBlue
    oRed = iRed
    oGreen = iGreen
    oBlue = iBlue
    DivFactor = 1
    L_Tol = DecimalPercent(TolFactor, 255)
    PixelRing Tx, TY, x, Y
    For i = 1 To 8
        If PixelSet And 2 ^ (i + 1) Then ' if it is a requested pixel Logical test for PixelValue being in PixelSet
            LongToRGB GetPixel(sour.hdc, Tx(i), TY(i)), iRed, iGreen, iBlue
            lngRed = lngRed + (oRed + L_Tol * iRed)
            lngGreen = lngGreen + (oGreen + L_Tol * iGreen)
            lngBlue = lngBlue + (oBlue + L_Tol * iBlue)
            DivFactor = DivFactor + 1
        End If
    Next i
    Brighten = RGB((lngRed / DivFactor), (lngGreen / DivFactor), (lngBlue / DivFactor))

End Function

Public Property Let Cancel(ByVal Cval As Boolean)

  'Write only (See Working for Read only partner)

    m_Cancel = Cval

End Property

Public Property Get ChangeWhileDrawing() As Boolean

    ChangeWhileDrawing = m_ChangeWhileDrawing

End Property

Public Property Let ChangeWhileDrawing(ByVal Value As Boolean)

    m_ChangeWhileDrawing = Value

End Property

Private Sub Class_Initialize()

  'set Default values

    m_GreyMode = False
    Red = True
    Blue = True
    green = True
    Invert = False
    Tolerance = 157 'Arivnder Sehmi used this and it seems to be a reasonable initial value
    ShowConversion = True
    AnyPixel = False

End Sub

Private Function ClockIntegerSubtraction(Val1 As Integer, val2 As Integer, limiter As Integer) As Integer

    ClockIntegerSubtraction = Val1 - val2
    Do While ClockIntegerSubtraction < 0
        ClockIntegerSubtraction = limiter + ClockIntegerSubtraction
    Loop

End Function

Private Function Darken(OwnCol As Long, x As Single, Y As Single, TolFactor As Integer) As Long

  'based on Chapter 5 'Visual Basic Power Toolkit' by R. Mansfield and E. Peroutsos 1955

  Dim iRed As Integer, iGreen As Integer, iBlue As Integer
  Dim lngRed As Long, lngGreen As Long, lngBlue As Long
  Dim oRed As Integer, oGreen As Integer, oBlue As Integer
  Dim DivFactor As Integer
  Dim i As Integer
  Dim Tx(8) As Single
  Dim TY(8) As Single
  Dim L_Tol As Single

    LongToRGB OwnCol, iRed, iGreen, iBlue
    lngRed = iRed
    lngGreen = iGreen
    lngBlue = iBlue
    DivFactor = 1
    L_Tol = DecimalPercent(TolFactor, 255)
    PixelRing Tx, TY, x, Y

    For i = 1 To 8
        If PixelSet And 2 ^ (i + 1) Then ' if it is a requested pixel Logical test for PixelValue being in PixelSet
            LongToRGB GetPixel(sour.hdc, Tx(i), TY(i)), iRed, iGreen, iBlue
            lngRed = lngRed + (oRed - (L_Tol * iRed))
            lngGreen = lngGreen + (oGreen - (L_Tol * iGreen))
            lngBlue = lngBlue + (oBlue - (L_Tol * iBlue))
            DivFactor = DivFactor + 1
        End If
    Next i

    If lngRed <= 0 Then
        lngRed = 0
    End If
    If lngGreen <= 0 Then
        lngGreen = 0
    End If
    If lngBlue <= 0 Then
        lngBlue = 0
    End If

    Darken = RGB((lngRed / DivFactor), (lngGreen / DivFactor), (lngBlue / DivFactor))

End Function

Private Function Darkenx(OwnCol As Long, TolFactor As Integer) As Long

  Dim RValue As Integer, GValue As Integer, BValue As Integer

    LongToRGB OwnCol, RValue, GValue, BValue

    If (RValue - TolFactor) >= 0 Then
    RValue = RValue - TolFactor
    End If
    If (GValue - TolFactor) >= 0 Then
    GValue = GValue - TolFactor
    End If
    If (BValue - TolFactor) >= 0 Then
    BValue = BValue - TolFactor
    End If
    Darkenx = RGB(RValue, GValue, BValue)

End Function

Private Function DecimalPercent(a, b) As Single

  'calculate DecimalPercentage given two numbers

  Dim c As Integer

    If a > b Then
        a = b
    End If
    If b = 0 Then
        DecimalPercent = 0
        Exit Function '>---> Bottom
    End If
    c = Int(a / b * 100)
    If c > 100 Then
        c = 100
    End If
    DecimalPercent = c / 100

End Function

Private Function Diffuse(OwnCol As Long, x As Single, Y As Single, TolFactor As Integer) As Long

  Dim iRed As Integer, iGreen As Integer, iBlue As Integer
  Dim lngRed As Long, lngGreen As Long, lngBlue As Long
  Dim DivFactor As Integer
  Dim i As Integer
  Dim Tx(8) As Single
  Dim TY(8) As Single

    Diffuse = OwnCol
    PixelRing Tx, TY, x, Y
    If Int(Rnd * m_Tolerance) > 255 / PixelSet Then
        i = Int(Rnd * 8) + 1
        If PixelSet And 2 ^ (i + 1) Then
            LongToRGB GetPixel(sour.hdc, Tx(i), TY(i)), iRed, iGreen, iBlue
            Diffuse = RGB(iRed, iGreen, iBlue)
        End If
    End If

End Function

Public Sub EdgeDetect() '(Optional Mode As EdgeMode = AllEdge)

  'Arivnder Sehmi provided all the thought for this
  'Arvinda's version used the equivalent of EdgeMode = LeftTopEdge
  'I just modifed it to work in my class and
  'created the WorkXXX and EdgeDetectTestStrong sub-routines to simplify the structure
  'EdgeDetectTestStrong uses EdgeMode to determine which edges to test

  Dim x As Single, Y  As Single
  Dim ImageSize As Single 'Used by WorkReport to allow reporting % done
  Dim SWidth As Single
  Dim SHeight As Single
  Dim TolFactor As Single
  Dim OwnCol As Long

    '  Dim ST As Single'DEBUG timing
    '  ST = Timer
    TolFactor = m_Tolerance ^ 3 ' This is faster but you can't change the factor as it runs
    ' To change it while running place inside the For Next structures
    WorkInitialize SWidth, SHeight, ImageSize
    For x = 1 To SWidth  'loop through the x-pixels
        For Y = 1 To SHeight  'loop through the y-pixels
            If m_ChangeWhileDrawing Then
                TolFactor = m_Tolerance ^ 3 ' This is slower as it recalculates Tolerence for each pixel
            End If                      'but lets you experiment while running
            OwnCol = GetPixel(sour.hdc, x, Y) 'get colour of target pixel
            If EdgeDetectTestStrong(OwnCol, x, Y, TolFactor) Then ', Mode) Then
                SetPixel Dest.hdc, x, Y, 0 ' plot pixel
            End If
        Next Y 'loop through the y-pixels

        WorkReport x, Y, ImageSize
        If m_Cancel Then
            Exit For '>---> Next
        End If
    Next x 'loop through the x-pixels
    WorkDone
    'Debug.Print Timer - ST

End Sub

Private Function EdgeDetectTest(OwnCol As Long, x As Single, Y As Single, TolFactor As Single) As Boolean

  'RG changed from Long to boolean

    EdgeDetectTest = Abs(OwnCol - GetPixel(sour.hdc, x, Y)) > TolFactor

End Function

Private Function EdgeDetectTestStrong(OwnCol As Long, x As Single, Y As Single, TolFactor As Single) As Boolean

  'Test surrounding pixels (1 to 8 depending on PixelSet)
  'IF m_AnyPixel = True then any queried pixels found to be True is enough to return True
  'else All queried pixels must match
  'NOTE If all pixels are set and m_AnyPixel=False then very few pixels will pass this test
  'Pixel No.   PixelValues
  '1 2 3       2    4   8
  '8   4       256      16
  '7 6 5       128  64  32
  'PixelSet is built by adding/subtracting PixelValues

  Dim T As Boolean ' hold a hit by EdgeDetectTest
  Dim i As Integer
  Dim Tx(8) As Single
  Dim TY(8) As Single

    PixelRing Tx, TY, x, Y

    For i = 1 To 8 'cycle through pixels
        If PixelSet And 2 ^ (i + 1) Then ' if it is a requested pixel Logical test for PixelValue being in PixelSet
            T = EdgeDetectTest(OwnCol, Tx(i), TY(i), TolFactor)
            If m_AnyPixel And T Then  'Any pixel will do so
                Exit For              'current pixel passed test '>---> Next
              ElseIf T = False Then    'M_AnyPixel=False'NOT M_ANYPIXEL...
                Exit For                 'current pixel failed test '>---> Next
            End If
        End If
    Next i

    EdgeDetectTestStrong = T
    If m_Invert Then 'Do invert if necessary
        EdgeDetectTestStrong = Not EdgeDetectTestStrong
    End If

End Function

Public Sub Experimental(SMth As Integer) '(Optional Mode As EdgeMode = AllEdge)

  'the experiments are based on Chapter 5 'Visual Basic Power Toolkit' by R. Mansfield and E. Peroutsos 1955
  'They coded for 256 colour bmps and palettes and manipulated pixels in complex arrays rather than directly
  'I have dropped the arrays and ignored the palette where it even exists.

  'You may recognize the basic shell of the routine is that of edge detect

  Dim x As Single, Y  As Single
  Dim ImageSize As Single 'Used by WorkReport to allow reporting % done
  Dim SWidth As Single
  Dim SHeight As Single
  Dim TolFactor As Integer
  Dim OwnCol As Long

    '  Dim ST As Single'DEBUG timing
    '  ST = Timer
    TolFactor = m_Tolerance  ' This is faster but you can't change the factor as it runs
    ' To change it while running place inside the For Next structures
    WorkInitialize SWidth, SHeight, ImageSize
    For x = 1 To SWidth  'loop through the x-pixels
        For Y = 1 To SHeight  'loop through the y-pixels
            If m_ChangeWhileDrawing Then
                TolFactor = m_Tolerance  ' This is slower as it recalculates Tolerence for each pixel
            End If                      'but lets you experiment while running
            OwnCol = GetPixel(sour.hdc, x, Y) 'get colour of target pixel
            Select Case SMth
              Case 0
                SetPixel Dest.hdc, x, Y, Smoother(OwnCol, x, Y, TolFactor) ' plot pixel
              Case 1
                SetPixel Dest.hdc, x, Y, Sharper(OwnCol, x, Y, TolFactor) ' plot pixel
              Case 2
                SetPixel Dest.hdc, x, Y, Sharper1(OwnCol, x, Y, TolFactor) ' plot pixel
              Case 3
                SetPixel Dest.hdc, x, Y, NeonCrayon(OwnCol, x, Y, TolFactor) ' plot pixel
              Case 4
                SetPixel Dest.hdc, x, Y, Diffuse(OwnCol, x, Y, TolFactor) ' plot pixel
              Case 5
                SetPixel Dest.hdc, x, Y, Brighten(OwnCol, x, Y, TolFactor) ' plot pixel
              Case 6
                SetPixel Dest.hdc, x, Y, Darken(OwnCol, x, Y, TolFactor) ' plot pixel
              Case 7
                SetPixel Dest.hdc, x, Y, FilterTo(OwnCol, fRed, TolFactor) ' plot pixel
              Case 8
                SetPixel Dest.hdc, x, Y, FilterTo(OwnCol, fGreen, TolFactor) ' plot pixel
              Case 9
                SetPixel Dest.hdc, x, Y, FilterTo(OwnCol, fBlue, TolFactor) ' plot pixel
              Case 10
                SetPixel Dest.hdc, x, Y, FilterTo(OwnCol, fYellow, TolFactor) ' plot pixel
              Case 11
                SetPixel Dest.hdc, x, Y, FilterTo(OwnCol, fPurple, TolFactor) ' plot pixel
              Case 12
                SetPixel Dest.hdc, x, Y, FilterTo(OwnCol, fAqua, TolFactor) ' plot pixel
              Case 13
                SetPixel Dest.hdc, x, Y, NegativeColourFilter1(OwnCol, TolFactor)  ' plot pixel
              Case 14
                SetPixel Dest.hdc, x, Y, NegativeColourFilter2(OwnCol, TolFactor)  ' plot pixel
              Case 15
                SetPixel Dest.hdc, x, Y, NegativeColourFilter3(OwnCol, TolFactor)  ' plot pixel
              Case 16
                 SetPixel Dest.hdc, x, Y, OwnCol  ' plot pixel
              Case 17
                SetPixel Dest.hdc, x, Y, Lighten(OwnCol, TolFactor)  ' plot pixel
              Case 18
                SetPixel Dest.hdc, x, Y, Darkenx(OwnCol, TolFactor)  ' plot pixel
            End Select

        Next Y 'loop through the y-pixels

        WorkReport x, Y, ImageSize
        If m_Cancel Then
            Exit For '>---> Next
        End If
    Next x 'loop through the x-pixels
    WorkDone
    'Debug.Print Timer - ST

End Sub

Private Function FilterTo(OwnCol As Long, cRGB As FilterTones, TolFactor As Integer) As Long

  'Based on work by Steven Paterson at
  'http://www.tolon.co.uk/ste/documents/maths.html

  'If you check his stuff you may find it hard to believe this is from the same universe
  'but a lot of what he does occurs in the routine that called this one or a coding equivalent anyway
  'although he got the math wrong for green so i used my own and I really don't know why he uses a
  'single For ..Next and resets the 2nd looping value rather than nesting two For..Next systems anyway..
  'I added FilterTones, the Yellow,Purple and Aqua filters
  'and the ability to vary the strength of the other colours
  'if TolFactor = 0 and cRGB = fRed or fGreen  or fBlue then you see Steven's colours
  
  Dim RValue As Integer, GValue As Integer, BValue As Integer

    LongToRGB OwnCol, RValue, GValue, BValue
    Select Case cRGB
      Case fRed
        GValue = MaximumInt(GValue, TolFactor)
        BValue = MaximumInt(BValue, TolFactor)
      Case fGreen
        RValue = MaximumInt(RValue, TolFactor)
        BValue = MaximumInt(BValue, TolFactor)
      Case fBlue
        RValue = MaximumInt(RValue, TolFactor)
        GValue = MaximumInt(GValue, TolFactor)
      Case fYellow
        RValue = MaximumInt(RValue, TolFactor)
      Case fPurple
        GValue = MaximumInt(GValue, TolFactor)
      Case fAqua
        BValue = MaximumInt(BValue, TolFactor)
    End Select
    FilterTo = RGB(RValue, GValue, BValue)

End Function

Public Property Get green() As Boolean

    green = m_Green

End Property

Public Property Let green(ByVal GVal As Boolean)

  'used in LineArtTest to include/exclude Green from grey conversion

    m_Green = PreventAllOff(GVal)

End Property

Private Function GreyLimiter(GreyC As Integer) As Integer

  'Creating the Integer variable Factor causes a
  'rounding off of the value of GreyC which when
  'used in the second line makes this work.

  'If you try to do it all in a single line this routine has no effect.
  'as the '/ m_Greylimit' and the '* m_Greylimit' cancel out.
  'This also works with Factor as Long but if you change Factor to a Single or Double this routine has no effect
  'as rounding fails
  'Using integer div '\' creates an effect very like Edge Detect Tolerance=0, Testing All Pixels, ANY sytle
  'If you know enough maths that explains this please let me know.

  Dim factor As Integer 'Long

    If m_GreyWeirdness Then
        ' This happens if you use the Huh?!? option in the Demo
        'Greys values <40 this produces some odd effects. Even values often (but not always) get greyed backgrounds
        '             <14 they can be very odd (especial even numbers).
        '             Higher values probably have subtler side effects. I would guess that the whites are actually very low greys.
        factor = GreyC \ m_Greylimit
      Else 'M_GREYWEIRDNESS = FALSE
        factor = GreyC / m_Greylimit
    End If
    GreyLimiter = factor * m_Greylimit

End Function

Public Property Get GreyMode() As Integer

    GreyMode = m_GreyMode

End Property

Public Property Let GreyMode(ByVal Value As Integer)

    m_GreyMode = Value

End Property

Private Function GreyPixel(longCol As Long) As Long

  'reduce any long colour to long grey scale

  Dim GreyValue As Integer, RValue As Integer, GValue As Integer, BValue As Integer

    LongToRGB longCol, RValue, GValue, BValue
    'NOTE If you add your own the 3 coefficents should sum to 1
    Select Case m_GreyMode
      Case 0
        GreyValue = RValue * 0.222 + GValue * 0.707 + BValue * 0.071 ' Found at PCS and several others
      Case 1
        GreyValue = RValue * 0.3 + GValue * 0.59 + BValue * 0.11     ' Mansfield & Petroutsos
      Case 2
        GreyValue = RValue * 1 / 3 + GValue * 1 / 3 + BValue * 1 / 3 'Exact thirds works properly only where R=G=B (sometimes called Pure Colours)
      Case 3
        GreyValue = BValue                                           'Blue Only
      Case 4
        GreyValue = RValue                                           'Red Only
      Case 5
        GreyValue = GValue                                           'Green Only
    End Select
    GreyValue = GreyLimiter(GreyValue)
    GreyPixel = RGB(GreyValue, GreyValue, GreyValue)
    If m_Invert Then
        GreyPixel = vbWhite - GreyPixel 'invert the colour if nessesarry
    End If

End Function

Public Property Let greys(ByVal Value As Integer)

    m_Greys = Value
    m_Greylimit = 255 / m_Greys

End Property

Public Property Get greys() As Integer

    greys = m_Greys

End Property

Public Sub GreyScaleImage()

  'I know there are faster ways to do this
  'but this fits with the general style of the class

  Dim GreyC As Long
  Dim x As Single, Y As Single, G_Tolerance As Long
  Dim ImageSize As Single 'Used by WorkReport to allow reporting % done
  Dim SWidth As Single
  Dim SHeight As Single

    WorkInitialize SWidth, SHeight, ImageSize
    If Not m_IgnoreTolerance Then
        G_Tolerance = GreyPixel(RGB(m_Tolerance, m_Tolerance, m_Tolerance))
      Else 'NOT NOT...
        G_Tolerance = vbBlack
    End If
    For x = 0 To SWidth
        For Y = 0 To SHeight ' cycle throght Image
            GreyC = GreyPixel(GetPixel(sour.hdc, x, Y))
            If GreyC < G_Tolerance Then
                GreyC = vbWhite
            End If

            SetPixel Dest.hdc, x, Y, GreyC   ' set pixel
        Next Y
        Dest.Refresh
        WorkReport x, Y, ImageSize
        If m_Cancel Then
            Exit For '>---> Next
        End If
    Next x
    WorkDone

End Sub

Public Property Let GreyWeirdness(ByVal Value As Boolean)

    m_GreyWeirdness = Value

End Property

Public Property Let IgnoreTolerance(ByVal Value As Boolean)

    m_IgnoreTolerance = Value

End Property

Public Property Get IgnoreTolerance() As Boolean

    IgnoreTolerance = m_IgnoreTolerance

End Property

Public Property Get Invert() As Boolean

    Invert = m_Invert

End Property

Public Property Let Invert(ByVal IVal As Boolean)

  'used in EdgeDetect and LineArt

    m_Invert = IVal

End Property

Private Function Lighten(OwnCol As Long, TolFactor As Integer) As Long

  Dim RValue As Integer, GValue As Integer, BValue As Integer

    LongToRGB OwnCol, RValue, GValue, BValue

    If (RValue + TolFactor) <= 255 Then
    RValue = RValue + TolFactor
    End If
    If (GValue + TolFactor) <= 255 Then
    GValue = GValue + TolFactor
    End If
    If (BValue + TolFactor) <= 255 Then
    BValue = BValue + TolFactor
    End If
    Lighten = RGB(RValue, GValue, BValue)

End Function

Private Function Lighten1(OwnCol As Long, TolFactor As Single) As Long

  Dim RGBc As Long

  Dim iBlue As Integer, iRed As Integer, iGreen As Integer

    iRed = OwnCol And &HFF
    iGreen = (OwnCol And &HFF00) / &H10000
    iBlue = (OwnCol And &HFF0000) / &H10000
    If (iRed + TolFactor) < 255 Then
    iRed = iRed + TolFactor
    End If
    If (iGreen + TolFactor) < 255 Then
    iGreen = iGreen + TolFactor
    End If
    If (iBlue + TolFactor) < 255 Then
    iBlue = iBlue + TolFactor
    End If
    Lighten1 = RGB(iRed, iGreen, iBlue)

End Function

Public Sub LineArt()

  'Arivnder Sehmi provided all the thought for this
  'I just modifed it to work in my class and
  'created the WorkXXX sub-routines and moved the Tolerance and Invert tests to
  'LineArtTest to speed up the structure

  Dim x As Single, Y As Single
  Dim ImageSize As Single ''Used by WorkReport to allow reporting % done
  Dim SWidth As Single
  Dim SHeight As Single

  Dim ST As Single 'DEBUG timing

    ST = Timer
    WorkInitialize SWidth, SHeight, ImageSize
    For x = 1 To SWidth 'loop through the x-pixels
        For Y = 1 To SHeight  'loop through the y-pixels
            If LineArtTest(GetPixel(sour.hdc, x, Y)) Then ' get Black or White for a coloured pixel
                SetPixel Dest.hdc, x, Y, 0  ' draw the pixel if it is black
            End If
        Next Y 'loop through the y-pixels
        WorkReport x, Y, ImageSize
        If m_Cancel Then
            Exit For '>---> Next
        End If
    Next x 'loop through the x-pixels
    WorkDone
    Debug.Print Timer - ST

End Sub

Private Function LineArtTest(longCol As Long) As Boolean

  'Arivnder Sehmi provided all the thought for this
  'RG substituted a simpler and slightly faster way of extracting RGB
  'only calculate colour if needed saves 15-20% in timing tests
  'and changed it to produce a Boolean return value

  Dim lr As Integer, lg As Integer, lb As Integer, GreyValue As Integer

    longCol = Abs(longCol) 'remove neg values which GetPixel can return and which crashes the RGB conversions
    If m_Red Then
        lr = longCol Mod 256
    End If
    If m_Green Then
        lg = (longCol And &HFF00FF00) / 256
    End If
    If m_Blue Then
        lb = (longCol And &HFF0000) / 65536
    End If
    '    gets a grey pixel value from RGB colour and compares to Tolerence
    LineArtTest = (lr * 0.222 + lg * 0.707 + lb * 0.071) < m_Tolerance
    'this is a more legit way to get grey values than / by number of colours used
    If m_Invert Then
        LineArtTest = Not LineArtTest 'invert the colours if nessesarry
    End If

End Function

Public Sub LongToRGB(longCol As Long, rRed As Integer, rGreen As Integer, rBlue As Integer)

  ' Get The Red, Blue And Green Values Of A Colour From The Long Value

    longCol = Abs(longCol) 'remove neg values which GetPixel can return and which crashes the RGB conversions
    rRed = longCol Mod 256 'Red
    rGreen = (longCol And &HFF00FF00) / 256 'Green
    rBlue = (longCol And &HFF0000) / 65536 'Blue

End Sub

Private Function MaximumInt(Val As Integer, Max As Integer) As Integer

    If Val > Max Then
        MaximumInt = Max
      Else 'NOT VAL...
        MaximumInt = Val
    End If

End Function

Private Function NegativeColourFilter2(OwnCol As Long, TolFactor As Integer) As Long

  'clock maths restricted range
  
  Dim RValue As Integer, GValue As Integer, BValue As Integer

    LongToRGB OwnCol, RValue, GValue, BValue
    RValue = ClockIntegerSubtraction(TolFactor, RValue, TolFactor)
    GValue = ClockIntegerSubtraction(TolFactor, GValue, TolFactor)
    BValue = ClockIntegerSubtraction(TolFactor, BValue, TolFactor)
    NegativeColourFilter2 = RGB(RValue, GValue, BValue)

      End Function

Private Function NegativeColourFilter1(OwnCol As Long, TolFactor As Integer) As Long

  'Abs maths
  
  Dim RValue As Integer, GValue As Integer, BValue As Integer

    LongToRGB OwnCol, RValue, GValue, BValue
    RValue = Abs(TolFactor - RValue)
    GValue = Abs(TolFactor - GValue)
    BValue = Abs(TolFactor - BValue)
    NegativeColourFilter1 = RGB(RValue, GValue, BValue)

      End Function

Private Function NegativeColourFilter3(OwnCol As Long, TolFactor As Integer) As Long

  'clock maths full range
  
  Dim RValue As Integer, GValue As Integer, BValue As Integer

    LongToRGB OwnCol, RValue, GValue, BValue
    RValue = ClockIntegerSubtraction(TolFactor, RValue, 255)
    GValue = ClockIntegerSubtraction(TolFactor, GValue, 255)
    BValue = ClockIntegerSubtraction(TolFactor, BValue, 255)
    NegativeColourFilter3 = RGB(RValue, GValue, BValue)

      End Function

'Private Function NegativeColourFilter4(ByVal OwnCol As Long, TolFactor As Integer) As Long
'
'  'Modified from
'  'http://www.tolon.co.uk/ste/documents/maths.html
'  'Copyright Steven Paterson
'  'this is supposed to be faster than other negative filters but it is not a true negative
'  'because at 255 tolerance it does not reverse it self
'  'but on 2nd and 3rd passes it does???
'
'  Dim RValue As Integer, GValue As Integer, BValue As Integer
''
''    OwnCol = OwnCol Xor &HFFFFFF
''
''RValue = OwnCol And &H100
''GValue = (OwnCol And &HFF00FF) / &H100 '
''BValue = (OwnCol And &HFF0000) / &H10000 '
''
'''    rRed = longCol Mod &H100 'Red
'''    rGreen = (longCol And &HFF00FF00) / &H100 'Green
'''    rBlue = (longCol And &HFF0000) / &H10000 'Blue
''
'
'
'    OwnCol = Abs(OwnCol) Xor vbWhite '&HFFFFFF
'    RValue = (OwnCol Mod &HFF)
'    GValue = (OwnCol And &HFF) / &H10000
'    BValue = (OwnCol And &HFF0000) / &H10000
'
'       RValue = ClockIntegerSubtraction(RValue * 2, RValue, TolFactor)
'       GValue = ClockIntegerSubtraction(GValue * 2, GValue, TolFactor)
'       BValue = ClockIntegerSubtraction(BValue * 2, BValue, TolFactor)
'    NegativeColourFilter4 = RGB(RValue, GValue, BValue)
'
'      End Function

Private Function NeonCrayon(OwnCol As Long, x As Single, Y As Single, TolFactor As Integer) As Long

  Dim iRed As Integer, iGreen As Integer, iBlue As Integer
  Dim oRed As Integer, oGreen As Integer, oBlue As Integer
  Dim nRed As Integer, nGreen As Integer, nBlue As Integer
  Dim tRed As Long, tGreen As Long, tBlue As Long
  Dim uRed As Long, uGreen As Long, uBlue As Long
  Dim DivFactor As Integer
  Dim i As Integer
  Dim Tx(8) As Single
  Dim TY(8) As Single

    LongToRGB OwnCol, oRed, oGreen, oBlue

    DivFactor = 1

    PixelRing Tx, TY, x, Y
    DivFactor = DivFactor + 1
    For i = 1 To 8
        If PixelSet And 2 ^ (i + 1) Then ' if it is a requested pixel Logical test for PixelValue being in PixelSet
            LongToRGB GetPixel(sour.hdc, Tx(i), TY(i)), iRed, iGreen, iBlue
            tRed = tRed + (oRed - iRed) ^ 2
            tGreen = tGreen + (oGreen - iGreen) ^ 2
            tBlue = tBlue + (oBlue - iBlue) ^ 2

            LongToRGB GetPixel(sour.hdc, Tx(i), TY(i)), iRed, iGreen, iBlue
            uRed = uRed + (oRed - iRed) ^ 2
            uGreen = uGreen + (oGreen - iGreen) ^ 2
            uBlue = uBlue + (oBlue - iBlue) ^ 2
            DivFactor = DivFactor + 1
        End If
    Next i

    nRed = Sqr((tRed + uRed) / DivFactor)
    If TolFactor / 3 > nRed Then
        nRed = 255
    End If
    nGreen = Sqr((tGreen + uGreen) / DivFactor)
    If TolFactor / 3 > nGreen Then
        nGreen = 255
    End If

    nBlue = Sqr((tBlue + uBlue) / DivFactor)
    If TolFactor / 3 > nBlue Then
        nBlue = 255
    End If

    NeonCrayon = RGB(nRed, nGreen, nBlue)
    'If NeonCrayon <= RGB(TolFactor, TolFactor, TolFactor) Then
    'NeonCrayon = vbWhite
    'End If

    If m_Invert Then 'Do invert if necessary
        NeonCrayon = vbWhite - NeonCrayon
    End If

End Function

Private Sub PixelRing(Tx As Variant, TY As Variant, x As Single, Y As Single)

  'Generate the array of marked positions for surrounding pixels

    Tx(1) = x - 1
    TY(1) = Y - 1 'LeftTop
    Tx(2) = x - 1
    TY(2) = Y     'Top
    Tx(3) = x - 1
    TY(3) = Y + 1 'RightTop
    Tx(4) = x
    TY(4) = Y + 1 'Right
    Tx(5) = x + 1
    TY(5) = Y + 1 'RightBottom
    Tx(6) = x + 1
    TY(6) = Y     'Bottom
    Tx(7) = x + 1
    TY(7) = Y - 1 'LeftBottom
    Tx(8) = x
    TY(8) = Y - 1 'Left

End Sub

Public Property Get PixelSet() As Integer

    PixelSet = m_PixelSet

End Property

Public Property Let PixelSet(ByVal Value As Integer)

  '' Place this in 'Private Sub chkPixel_Click(Index As Integer)'
  ''to set Value
  ''    Dim val As Integer
  ''    val = 2 ^ (Index + 1)
  ''    LA.PixelSet = LA.PixelSet + IIf(chkPixel(Index).Value = vbChecked, val, -val)
  ''

    m_PixelSet = Value

End Property

Private Function PreventAllOff(TestValue As Boolean) As Boolean

  'make sure at least one colour band is being used for LineArtTest
  ' See KeepOneColorOn in the demo program for keeping the
  ' forms checkboxes in sync with this

    If m_Blue = False And m_Red = False And m_Green = False Then
        PreventAllOff = True
      Else 'NOT PREVENTALLOFF...'NOT M_BLUE...
        PreventAllOff = TestValue
    End If

End Function

Public Property Get Red() As Boolean

    Red = m_Red

End Property

Public Property Let Red(ByVal RVal As Boolean)

  'used in LineArtTest to include/exclude Red from grey conversion

    m_Red = PreventAllOff(RVal)

End Property

Private Function Sharpener(x As Single, Y As Single, Tol As Long) As Single

  Dim i As Integer, nextcol As Long, OwnCol As Long
  Dim Tx(8) As Single
  Dim TY(8) As Single

    PixelRing Tx, TY, x, Y
    OwnCol = GetPixel(sour.hdc, x, Y)
    Sharpener = OwnCol
    For i = 1 To 8 'cycle through pixels
        If PixelSet And 2 ^ (i + 1) Then ' if it is a requested pixel Logical test for PixelValue being in PixelSet
            nextcol = GetPixel(sour.hdc, Tx(i), TY(i))
            If nextcol < Tol Then
                Sharpener = Sharpener + 0.5 * (OwnCol - nextcol)
            End If
        End If
    Next i
    If m_Invert Then
        Sharpener = vbWhite - Sharpener 'invert the colour if nessesarry
    End If

End Function

Public Sub SharpenImage()

  'I know there are faster ways to do this
  'but this fits with the general style of the class

  Dim GreyC As Long
  Dim x As Single, Y As Single, G_Tolerance As Long
  Dim ImageSize As Single 'Used by WorkReport to allow reporting % done
  Dim SWidth As Single
  Dim SHeight As Single

    WorkInitialize SWidth, SHeight, ImageSize
    '    If Not m_IgnoreTolerance Then
    '    G_Tolerance = GreyPixel(RGB(m_Tolerance, m_Tolerance, m_Tolerance))
    '    Else
    '    End If
    For x = 0 To SWidth
        For Y = 0 To SHeight ' cycle throght Image
            G_Tolerance = vbBlack - RGB(m_Tolerance, m_Tolerance, m_Tolerance)

            GreyC = Sharpener(x, Y, G_Tolerance)
            '            If GreyC < G_Tolerance Then
            '                GreyC = vbWhite
            '            End If

            SetPixel Dest.hdc, x, Y, GreyC   ' set pixel
        Next Y
        Dest.Refresh
        WorkReport x, Y, ImageSize
        If m_Cancel Then
            Exit For '>---> Next
        End If
    Next x
    WorkDone

End Sub

Private Function Sharper(OwnCol As Long, x As Single, Y As Single, TolFactor As Integer) As Long

  Dim lngRed As Long, lngGreen As Long, lngBlue As Long
  Dim iRed As Integer, iGreen As Integer, iBlue As Integer
  Dim oRed As Integer, oGreen As Integer, oBlue As Integer
  Dim DivFactor As Integer

  Dim i As Integer
  Dim Tx(8) As Single
  Dim TY(8) As Single

    LongToRGB OwnCol, iRed, iGreen, iBlue

    oRed = iRed
    oGreen = iGreen
    oBlue = iBlue
    lngRed = iRed
    lngGreen = iGreen
    lngBlue = iBlue
    DivFactor = 1
    PixelRing Tx, TY, x, Y
    For i = 1 To 8
        If PixelSet And 2 ^ (i + 1) Then ' if it is a requested pixel Logical test for PixelValue being in PixelSet
            'If i = 1 Or i = 8 or iThen
            LongToRGB GetPixel(sour.hdc, Tx(i), TY(i)), iRed, iGreen, iBlue
            lngRed = lngRed - iRed
            lngGreen = lngGreen - iGreen
            lngBlue = lngBlue - iBlue

            DivFactor = DivFactor + 1
        End If
    Next i
    Sharper = RGB(Abs(oRed - lngRed / DivFactor), Abs(oGreen - lngGreen / DivFactor), Abs(oBlue - lngBlue / DivFactor))

End Function

Private Function Sharper1(OwnCol As Long, x As Single, Y As Single, TolFactor As Integer) As Long

  Dim lngRed As Long, lngGreen As Long, lngBlue As Long
  Dim iRed As Integer, iGreen As Integer, iBlue As Integer
  Dim oRed As Integer, oGreen As Integer, oBlue As Integer
  Dim i As Integer, L_Tol As Single
  Dim Tx(8) As Single
  Dim TY(8) As Single

    LongToRGB OwnCol, iRed, iGreen, iBlue
    oRed = iRed
    oGreen = iGreen
    oBlue = iBlue
    L_Tol = DecimalPercent(TolFactor, 255) / 100
    PixelRing Tx, TY, x, Y
    For i = 1 To 8
        If PixelSet And 2 ^ (i + 1) Then ' if it is a requested pixel Logical test for PixelValue being in PixelSet
            'If i = 1 Or i = 8 or iThen
            LongToRGB GetPixel(sour.hdc, Tx(i), TY(i)), iRed, iGreen, iBlue
            lngRed = lngRed + L_Tol * (oRed - iRed)
            lngGreen = lngGreen + L_Tol * (oGreen - iGreen)
            lngBlue = lngBlue + L_Tol * (oBlue - iBlue)
        End If
    Next i
    Sharper1 = RGB(Abs(oRed + lngRed), Abs(oGreen + lngGreen), Abs(oBlue + lngBlue))

End Function

Public Property Get ShowConversion() As Boolean

    ShowConversion = m_ShowConversion

End Property

Public Property Let ShowConversion(ByVal SVal As Boolean)

  'used in WorkReport

    m_ShowConversion = SVal

End Property

Private Function Smoother(OwnCol As Long, x As Single, Y As Single, TolFactor As Integer) As Long

  'based on Chapter 5 'Visual Basic Power Toolkit' by R. Mansfield and E. Peroutsos 1955

  Dim iRed As Integer, iGreen As Integer, iBlue As Integer
  Dim lngRed As Long, lngGreen As Long, lngBlue As Long
  Dim DivFactor As Integer
  Dim i As Integer
  Dim Tx(8) As Single
  Dim TY(8) As Single, L_Tol As Single

    LongToRGB OwnCol, iRed, iGreen, iBlue
    lngRed = iRed
    lngGreen = iGreen
    lngBlue = iBlue
    DivFactor = 1
    PixelRing Tx, TY, x, Y
    For i = 1 To 8
        If PixelSet And 2 ^ (i + 1) Then ' if it is a requested pixel Logical test for PixelValue being in PixelSet
            LongToRGB GetPixel(sour.hdc, Tx(i), TY(i)), iRed, iGreen, iBlue
            lngRed = lngRed + iRed
            lngGreen = lngGreen + iGreen
            lngBlue = lngBlue + iBlue
            DivFactor = DivFactor + 1
        End If
    Next i
    Smoother = RGB(lngRed / DivFactor, lngGreen / DivFactor, lngBlue / DivFactor)

End Function

Public Property Get Tolerance() As Long

    Tolerance = m_Tolerance

End Property

Public Property Let Tolerance(ByVal Tval As Long)

  'used in EdgeDetect and LineArt

    m_Tolerance = Tval

End Property

Private Sub WorkDone()

  'RG
  ' encapsulate common parts of finishing EdgeDetect and LineArt Routines

    Dest.Refresh
    m_Working = False
    DecimalPercentBox.Caption = ""

End Sub

Public Property Get Working() As Boolean

  'Read only (See Cancel for Write only partner)
  'Reports to program if class is working

    Working = m_Working

End Property

Private Sub WorkInitialize(SWidth As Single, SHeight As Single, ImageSize As Single)

  'RG
  ' encapsulate common parts of start of EdgeDetect and LineArt Routines

    m_Cancel = False
    m_Working = True
    With sour
        SHeight = .Height / Screen.TwipsPerPixelY
        SWidth = .Width / Screen.TwipsPerPixelX
        ImageSize = SWidth * SHeight
    End With 'SOUR
    Dest.Cls
    Dest.Refresh

End Sub

Private Sub WorkReport(x As Single, Y As Single, ImageSize As Single)

  'RG
  ' encapsulate common parts of running of EdgeDetect and LineArt Routines

    DoEvents
    If ShowConversion Then
        Dest.Refresh
    End If
    If DecimalPercentBoxExists Then
        DecimalPercentBox.Caption = Int(((x * Y) / ImageSize) * 100) & "%" 'calculate the DecimalPercent done.
    End If

End Sub

':) Ulli's VB Code Formatter V2.13.6 (27/12/2002 8:38:33 PM) 149 + 1132 = 1281 Lines
