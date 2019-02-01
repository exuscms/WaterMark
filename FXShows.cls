VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FXShows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'about 1 TO 10 MB of Memory Needed, DEPENDING ON PHOTO SIZE
Private BitsOf1 As BITMAPINFO 'The Bits Of the First Picture Box
Private BitsOf2 As BITMAPINFO 'The Bits Of the Second Picture Box
Private DestBits As BITMAPINFO 'The Final Bits

'local variable(s) to hold property value(s)
Private mvarSourceDC1 As Object 'local copy
Private mvarSourceDC2 As Object 'local copy
Private mvarDestDC1 As Object 'local copy
Public CurrectEffect As Effects

Public Sub ExecuteEffect(Optional ByVal sPercentage As Single = 50)

'Process Selected Effect
Select Case CurrectEffect

Case None
    exeNone BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case Fade
    exeFade BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case Dissolove
    exeDissolove BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Horizontal Bars]
    exeHorizontalBars BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Vertical Bars]
    exeVerticalBars BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Box IN]
    exeBoxIN BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Box OUT]
    exeBoxOUT BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Pixelate OUT IN]
    exePixelateOUTIN BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Chess Boxes]
    exeChessBoxes BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case Diffuse
    exeDiffuse BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Brightness IN OUT]
    exeBrightnessINOUT BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Blackness IN OUT]
    exeBlacknessINOUT BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case TV
    exeTV BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Random Horizontal Lines]
    exeRandomHorizontalLines BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Random Vertical Lines]
    exeRandomVerticalLines BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Wip Horizontal]
    exeWipeHorizontal BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Wipe Vertical]
    exeWipeVertical BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Slide Up]
    exeSlideUp BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Slide Down]
    exeSlideDown BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Move Left - Right]
    exeMoveLeftRight BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Move Right - Left]
    exeMoveRightLeft BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Move Down - Up]
    exeMoveDownUp BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

Case [Move Up - Down]
    exeMoveUpDown BitsOf1, BitsOf2, DestBits, CInt(sPercentage)

End Select

'SetPicture from points (bits)
SetBITMAPINFO mvarDestDC1, DestBits
End Sub

Public Sub PrepareForNew()
'GETTING POINTS OF THE FIRST PICTURE
GetBITMAPINFO mvarSourceDC1, BitsOf1
'GETTING POINTS OF THE SECOND PICTURE
GetBITMAPINFO mvarSourceDC2, BitsOf2
'SETTING THE DISTINATION POINTS AS THE POINTS OF THE FIRST PICTURE
DestBits = BitsOf1
End Sub


Public Property Set DestDC1(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DestDC1 = Form1
    Set mvarDestDC1 = vData
    DestDC = vData.hdc
End Property


Public Property Get DestDC1() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DestDC1
    Set DestDC1 = mvarDestDC1
End Property



Public Property Set SourceDC2(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SourceDC2 = Form1
    Set mvarSourceDC2 = vData
    SrcDC2 = vData.hdc
End Property


Public Property Get SourceDC2() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SourceDC2
    Set SourceDC2 = mvarSourceDC2
End Property



Public Property Set SourceDC1(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SourceDC1 = Form1
    Set mvarSourceDC1 = vData
    SrcDC1 = vData.hdc
End Property


Public Property Get SourceDC1() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SourceDC1
    Set SourceDC1 = mvarSourceDC1
End Property



