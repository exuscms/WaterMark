VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFramedScrollPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'jan 2 2003
'Updated SetScroll added

'some of this is copyright Min Thant Sin
'based on code found at
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=41164&lngWId=1
'Modifications Copyright 2002 Roger Gilchrist rojagilkrist@hotmail.com
'converted to a class demonstrating use of WithEvents and created the centre small image stuff and interlocking.
'
'UPDATE Class now assumes that there is a Frame containing two PictureBoxes, a vertical ScrollBar
'         and a Horizontal ScrollBar.
'One of the PictureBoxes must be contained by the other.
'This contained Picturebox is the only control you need to tell the class about
'The class automatically identifies all the other controls.

'DESCRIPTION
'This class allows you to use a PictureBox containing a picture larger than the view area
'so that you can scroll around the image using Scroll bars, Mouse or Arrow keys
'Pictures smaller than the viewing area are centred in the viewing area and scrolling is disabled.
'A Picture can be scrollable in only one axis if it fits on the other axis.
'The class also allows you to co-ordinate a second Scrollable picture box so that its layout and size match the first one.
'
'REQUIRES: PictureBox contained within another PictureBox, a VScrollBar and a HScrollBar.
'          All these are contained by a Frame.
'          Set the Frame to the size you want. Everything else will be sized by the class.
'
'Place in Declaration section of a Form|Module
'
'   Private|Public ScrollPic As New ClsScrollPicture
'
'Place in Form_Load|Sub Main
'
'       ScrollPic.AssignControls PicBox
' Where PicBox is the PictureBox you will be loading pictures to.
'
'[OPTIONAL] Place this just below AssignControls call, if you have two PictureBoxes which are linked.
'           It allows you to interlock the scrollbars of the two PictureBox systems
'           and make sure secondry system is identical in appearance to Primary ScrollablePicture
'       ScrollPic2.AssignControls PicBox2
'       ScrollPic.AssignInterlockScrolls Vscroll2, HScroll2
'       ScrollPic2.AssignInterlockScrolls Vscroll, HScroll
'       ScrollPic2.ResizeTo PicBox

'
'Place just after any picture loading code
'      [COMPULSORY] ScrollPic.NewPicture [Optional Caption]
'      [OPTIONAL] ScrollPic2.ResizeTo PicBox
'                      use only if you need to set a second PictureBox to the same size as
'                      the ScrollPic PictureBox (See LineArt Demo for examples)
''Sets the values for scrollbars and size of movements relative to the image size
'
'PROPERTIES these have Default values you probably don't need to change but if you want you can
'           change Defaults in Class_Initialize
''Caption        Only useful if there is a Frame around the PictureBoxes
''ContainerColor DEF =RGB(127,127,127) colour of container picturebox, visible as a frame if image is smaller than Viewing area.
''Scroll_Percent DEF = 0.1 set LargeChange value for scroll bars to a percentage of image size
''Move_Speed     DEF = 100 size of steps used by arrow keys
''ContainsImage  (experimental) tests image pixals to determine if ther is an image in the PictureBox
''FrameOffset    Def = 80  set standard offset of controls from edge of Frame if it has BorderStyle=1 else = 0
''PixelSampleRate Def=1 set percent of diagonal pixel to test in PixelSampler
''ScrollSize     Def = 150 set standard width for VScroll and Height for HScroll

'That's it!
'
'Well not really but the following are internal to the class and never called directly.
'
'Private Routines
'DoScroll           Does calculations for scroll actions. Blocks scroll if not needed
'KeepInBounds       Keeps a value between Min and Max values
'SetScrolls         find Scrollbars in same Frame as picturebox (use to get both own and interlocking scrolls
'PositionElements   Set pictureBoxes and Scroll bars in correct places on Frame
'PixelSampler  (experimental) Support routine for ContainsImage

'WithEvents Routines
'''Not all Events are used/needed for the class to work only the following are supplied
'Because the class uses WithEvents all the behaviour is automatically executed
''If you need to create any other activities in these Events remember that the On_Form_Event
''fires before the matching Class_Event. So On_Form cannot use Class generated response data

''M_MovingPic_GotFocus  Turn on KeyPreview for m_PicForm_KeyDown
''M_MovingPic_LostFocus Turn off KeyPreview for m_PicForm_KeyDown
''M_MovingPic_MouseDown Set Start point for Mouse scrolling
''M_MovingPic_MouseMove Move m_MovingPic with mouse
''m_PicForm_KeyDown     Read Arrow keys if Picture has focus
''VScrl_Change          Move m_MovingPic with scrollbar
''VScrl_Scroll          Move m_MovingPic with scrollbar (Updates while you hold slider button)
''HScrl_Change
''HScrl_Scroll

''OTHER NOTES
''The TabStop setting for all controls except m_MovingPic is set to False by the class
''because you only need one TabStop for the FrameScrollingPicture. Be aware of this
''if you want to change TabOrder on your form.

Private m_ScrollSize As Integer
Private m_FrameOffset As Integer
Private m_Scroll_Percent As Single
Private m_Move_Speed As Integer
Private LastMouseX As Single  'self-explanatory
Private LastMouseY As Single  '           "
Private tmpLeft As Integer
Private tmpTop As Integer
Private XDiff As Integer      'width difference
Private YDiff As Integer      'height difference
Private HasInterlock As Boolean
Private InterlockVScrl As VScrollBar
Attribute InterlockVScrl.VB_VarHelpID = -1
Private InterlockHScrl As HScrollBar
Attribute InterlockHScrl.VB_VarHelpID = -1
Private HScrollable As Boolean ' If picture is shorter than Viewing Height don't HScroll
Private VScrollable As Boolean ' If picture is narrower than Viewiewing Width don't VScroll
Private WithEvents M_MovingPic As PictureBox
Attribute M_MovingPic.VB_VarHelpID = -1
'This is the picturebox you load to/draw on whatever
Private m_ContainerPic As PictureBox
Private m_ContainerFrame As Frame
Private m_Caption As String
Private m_ContainerBackColor As Long
Private m_ResizingLock As Boolean ' Stops resetting Scroll values from firing Scroll_Change events
Private m_PixelSampleRate As Integer
'M_ContainerPic is used to set the on-screen size of the picture
'M_ContainerPic does not need to be assigned as the
'Sub AssignControls uses M_MovingPic.Container to get it
Private WithEvents VScrl As VScrollBar
Attribute VScrl.VB_VarHelpID = -1
Private WithEvents HScrl As HScrollBar
Attribute HScrl.VB_VarHelpID = -1
Private WithEvents m_PicForm As Form
Attribute m_PicForm.VB_VarHelpID = -1
Private InitialPixelValues As Single
Private m_ForceContains As Boolean
'm_PicForm is the form on which the pictureboxes are located
'It is needed to allow KeyPreview to be set and the Arrow keys to scroll
'm_PicForm does not need to be assigned as the
'Sub AssignControls uses M_MovingPic.Parent to get it
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long

Public Sub AssignControls(MovePic As PictureBox, Optional Capt As String)
'Copyright 2002 Roger Gilchrist
  'Set the Various controls that the class manipulates
  'NOTE M_ContainerPic, M_ContainerFrame, m_PicForm, VScrl and HScrl must exist but are not assigned
  'using .Container, .Parent and m_PicForm.Controls the class finds them automatically.

  Dim ErrMsg As String        'Generic error if any controls are missing.

    ErrMsg = "This class has the following requirements:" & vbNewLine & _
             "1. The PictureBox sent to AssignControls has to be contained within another PictureBox" & vbNewLine & _
             "2. The containing PictureBox must be contained in a Frame." & vbNewLine & _
             "3. There are two Scrollbars, one Horizontal and one Vertical contained in the Frame." & vbNewLine & vbNewLine & _
             "Failure to fulfill all of these requirements ends the program." & vbNewLine & _
             "You have failed to fullfil requirement no. "

    Set M_MovingPic = MovePic

    If Not TypeOf M_MovingPic.Container Is PictureBox Then 'Test Requirement 1
        MsgBox ErrMsg & "1", vbCritical, "ClsFramedScrollPicture"
        End
    End If
    If Not TypeOf M_MovingPic.Container.Container Is Frame Then 'Test Requirement 2
        MsgBox ErrMsg & "2", vbCritical, "ClsFramedScrollPicture"
        End
    End If

    With M_MovingPic
        .AutoSize = True
        .ZOrder
        Set m_ContainerPic = .Container
        m_ContainerPic.TabStop = False
        Set m_ContainerFrame = m_ContainerPic.Container

        m_ContainerFrame = Capt
        Set m_PicForm = .Parent
    End With 'M_MOVINGPIC
    SetScrolls M_MovingPic, ErrMsg
    PositionElements '   Place all elements
    M_MovingPic.Move 0, 0, m_ContainerPic.Width, m_ContainerPic.Height
    SetMoveValues

End Sub

Public Sub AssignInterlockScrolls(pic As PictureBox)
'Copyright 2002 Roger Gilchrist
  'Optional Routine
  'Allows you to create pairs of scrollable PictureBoxes which are interlocked
  '(a move in one is reflected in the other) by linking their respective scrollbars

  Dim EMsg As String

    EMsg = "Interlock controls not found."
    SetScrolls pic, EMsg
    If Len(EMsg) Then
        MsgBox EMsg, vbCritical, "ClsFrameScrollingPicture"
      Else 'LEN(EMSG) = FALSE
        HasInterlock = True '    let rest of class know if this has been activated
    End If

End Sub

Public Property Let Caption(ByVal Value As String)

  'This does nothing unless Frame.BorderStyle=1

    m_Caption = Value
    m_ContainerFrame.Caption = m_Caption

End Property

Public Property Get Caption() As String

    Caption = m_Caption

End Property

Private Sub Class_Initialize()

  'Set Default Values

    m_PixelSampleRate = 10
    m_FrameOffset = 80
    m_ScrollSize = 150
    m_Move_Speed = 100
    m_Scroll_Percent = 0.1
    ContainerColor = RGB(128, 128, 128)

End Sub

Public Property Let ContainerColor(ByVal Value As Long)

  'set the colour of the surrounding frame if image is smaller than viewing area
  'Def= MidGrey RGB(128, 128, 128)

    m_ContainerBackColor = Value

End Property

Public Property Get ContainerColor() As Long

    ContainerColor = m_ContainerBackColor

End Property

Public Property Get ContainsImage() As Boolean
'Copyright 2002 Roger Gilchrist
  'test whether pixels have changed, so image exists
  'This routine depends on sampling the PictureBox so can fail
  '(See Let for details on how to over-ride the test)
  '
  'If you know a simpler way of determining whether there is an image in a PictureBox
  'without setting a variable as you load or draw to the PictureBox let me know.

    If m_ForceContains Then 'the Property Let allows you to over-ride the setting
        ContainsImage = True
      Else 'M_FORCECONTAINS = FALSE
        If Not M_MovingPic Is Nothing Then
            ContainsImage = InitialPixelValues <> PixelSampler
        End If
    End If
    m_ForceContains = False

End Property

Public Property Let ContainsImage(Value As Boolean)

  'Allows you to force class to believe it contains an image
  'So you can override the internal test if necessary
  'Normally you would never do this but the test is falible
  'so you may need to force it (See Get for more details)
  'You could also try raising the PixelSampleRate although this can slow the program dramatically

    m_ForceContains = True

End Property

Private Sub DoScroll()

  'Make sure M_MovingPic doesn't go out of boundaries
  'Does calculations for scroll actions. does nothing if scroll not required
  'Note negative inversion of XDiff, YDiff values for KeepInBounds

    If HScrollable Then
        KeepInBounds -XDiff, tmpLeft, 0
        HScrl.Value = Abs(tmpLeft)
    End If
    If VScrollable Then
        KeepInBounds -YDiff, tmpTop, 0
        VScrl.Value = Abs(tmpTop)
    End If

End Sub

Public Property Get FrameOffSet() As Long

    FrameOffSet = m_FrameOffset

End Property

Public Property Let FrameOffSet(ByVal Value As Long)

  'redesign distance betwwen sub controls and frame if Frame.BorderStyle=1 Def = 80

    m_FrameOffset = Value

End Property

Private Sub HScrl_Change()
'Copyright 2002 Roger Gilchrist
  'WithEvents causes this to be called

    If Not m_ResizingLock Then
        M_MovingPic.Left = (-HScrl.Value)
        tmpLeft = M_MovingPic.Left
        'move interlocked bar if it exists
        If HasInterlock Then
            InterlockHScrl.Value = HScrl.Value
        End If
    End If

End Sub

Private Sub HScrl_Scroll()

  'WithEvents causes this to be called

    HScrl_Change

End Sub

Private Sub KeepInBounds(Min As Integer, Val As Integer, Max As Integer)

  'keeps Val between Min and Max

    If Val >= Max Then
        Val = Max
    End If
    If Val <= Min Then
        Val = Min
    End If

End Sub

Private Sub M_MovingPic_GotFocus()

    m_PicForm.KeyPreview = True

End Sub

Private Sub M_MovingPic_LostFocus()

    m_PicForm.KeyPreview = False

End Sub

Private Sub M_MovingPic_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) '

  'WithEvents causes this to be called
  'sets initial postion of mouse grag for mouse scrolling

    If Button = vbLeftButton Then
        LastMouseX = x
        LastMouseY = Y
    End If

End Sub

Private Sub M_MovingPic_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

  'WithEvents causes this to be called
  'Drags the image
  'NOTE M_MovingPic_MouseUp is not needed as only The LastMouseX|Y data
  'needs to be dealt with and MouseDown does that

    If Button = vbLeftButton Then
        tmpLeft = M_MovingPic.Left + IIf(HScrollable, (x - LastMouseX), 0)
        tmpTop = M_MovingPic.Top + IIf(VScrollable, (Y - LastMouseY), 0)
        DoScroll
    End If

End Sub

Private Sub m_PicForm_KeyDown(KeyCode As Integer, Shift As Integer)

  'WithEvents causes this to be called

    Select Case KeyCode
      Case vbKeyLeft, vbKeyRight
        tmpLeft = tmpLeft + IIf(HScrollable, IIf(vbKeyRight = KeyCode, -m_Move_Speed, m_Move_Speed), 0)
      Case vbKeyUp, vbKeyDown
        tmpTop = tmpTop + IIf(VScrollable, IIf(vbKeyDown = KeyCode, -m_Move_Speed, m_Move_Speed), 0)
    End Select
    DoScroll

End Sub

Public Property Let Move_Speed(ByVal Value As Integer)

  'Allows you to Let size of steps used by arrow keys DEF = 100

    m_Move_Speed = Value

End Property

Public Property Get Move_Speed() As Integer

  'Allows yout Get size of steps used by arrow keys DEF = 100

    Move_Speed = m_Move_Speed

End Property

Public Sub NewPicture(Optional CaptionStr As String)

  'This is really only to set the m_ContainsPicture if you load a picture

    If Len(CaptionStr) Then
        Caption = CaptionStr
    End If
    SetMoveValues

End Sub

Private Function PixelSampler() As Single
'Copyright 2002 Roger Gilchrist
  'Sample the diagonals approximately every m_PixelSampleRate% of length
  'This is used to test if a picture has been drawn in the PictureBox
  'by comparing with the value tested when the picturebox was set.
  'slows things down a bit and may fail if the image happens not
  'to fall on any of the tested pixels but its's better than nothing
  'for example of failure try 'Edge Detection' Top and Left pixels at Tolerencee 202 and 203
  ''using the supplied LineArtDemo.jpg and the default m_PixelSampleRate = 10
  'at 202 this routine returns True but
  'at 203 (or above), even though there are pixels in the image, none of them are noticed by PixelSampler.

  Dim i As Long, j As Long, TPixel As Long, SRate As Single

    SRate = M_MovingPic.Width * m_PixelSampleRate / 100
    For i = 1 To M_MovingPic.Width Step SRate
        For j = 1 To M_MovingPic.Height Step SRate
            TPixel = M_MovingPic.Point(i, j) + M_MovingPic.Point(M_MovingPic.Width - i, M_MovingPic.Height - j) 'GetPixel(M_MovingPic.hdc, i, j) '+ GetPixel(M_MovingPic.hdc, j, i)
            If j Mod 2 Then
                PixelSampler = PixelSampler + TPixel
              Else 'NOT J...
                PixelSampler = PixelSampler - TPixel
            End If
        Next j
    Next i

End Function

Public Property Let PixelSampleRate(ByVal Value As Integer)

  'This improves the probability that PixelSampler will detect a change but
  'causes the program to slow down dramatically as the number rises.
  'max = 20 is arbitary but probably as high as you would want to go

    KeepInBounds 1, Value, 20
    m_PixelSampleRate = Value

End Property

Public Property Get PixelSampleRate() As Integer

    PixelSampleRate = m_PixelSampleRate

End Property

Private Sub PositionElements()
'Copyright 2002 Roger Gilchrist
  'This routine places all the controls in the correct positions relative to the Frame

  Dim GeneralOffSet As Long ' hold the offset size for left, right and bottom edges of Frame
  Dim TopOffset As Long 'Deal with the presence/absence of a Caption and fact that Topoffset is twice GeneralOffSet

    With m_ContainerFrame
        'If there is a Frame.Border then Offset from Frame edge to show border else set to edge
        GeneralOffSet = IIf(.BorderStyle = 1, m_FrameOffset, 0)
        'If there is a Caption OfSet 2.5* GeneralOffSet else 2*; Ignore Captions on Borderless Frames
        TopOffset = IIf(.BorderStyle = 1, GeneralOffSet * IIf(Len(.Caption), 2.5, 2), 0)
        'Draw HScroll
        HScrl.Top = .Height - m_ScrollSize - GeneralOffSet
        HScrl.Left = GeneralOffSet
        HScrl.Height = m_ScrollSize
        HScrl.Width = .Width - m_ScrollSize - GeneralOffSet * 2
        HScrl.TabStop = 0
        'Draw VScroll
        VScrl.Top = TopOffset
        VScrl.Left = .Width - m_ScrollSize - GeneralOffSet
        VScrl.Height = HScrl.Top - TopOffset
        VScrl.Width = m_ScrollSize
        VScrl.TabStop = 0
        'Draw Container
        m_ContainerPic.Left = GeneralOffSet
        m_ContainerPic.Top = TopOffset
        m_ContainerPic.Height = HScrl.Top - TopOffset
        m_ContainerPic.Width = VScrl.Left - GeneralOffSet
        m_ContainerPic.BackColor = m_ContainerBackColor

    End With 'M_CONTAINERFRAME

End Sub

Public Sub ResizeTo(pic As PictureBox)
'Copyright 2002 Roger Gilchrist
  'This routine allows you to set a ScrollPicture system to the same size as another one.
  'UPDATED: Now only needs the PicTureBox from the other class implemetation to operate

    If Not (m_ContainerFrame Is Nothing) Then 'Safety Trap
        With m_ContainerFrame 'Match Frame Sizes
            .Width = pic.Container.Container.Width
            .Height = pic.Container.Container.Height
        End With 'M_CONTAINERFRAME
        PositionElements ' set up all other controls
        With M_MovingPic ' force secondary ScrollingPicture size to primary size
            .Picture = LoadPicture() 'Clear existing pic
            .Height = pic.Height
            .Width = pic.Width
        End With 'M_MOVINGPIC
        SetMoveValues
      Else 'NOT NOT...
        MsgBox "You need to call AssignControls before calling ResizeTo", vbCritical, "ClsFramedScrollingPicture"
        End
    End If
    
End Sub

Public Property Let Scroll_Percent(ByVal Value As Single)

  'Allows you to Let LargeChange value for scroll bars to a percentage of image size DEF = 0.1

    m_Scroll_Percent = Value

End Property

Public Property Get Scroll_Percent() As Single

  'Allows you to Get LargeChange value for scroll bars to a percentage of image size DEF = 0.1

    Scroll_Percent = m_Scroll_Percent

End Property

Public Property Let ScrollSize(ByVal Value As Long)

  'redesign scroll bar size Def=150

    m_ScrollSize = Value

End Property

Public Property Get ScrollSize() As Long

    ScrollSize = m_ScrollSize

End Property

Public Sub SetMoveValues()
'Copyright 2002 Roger Gilchrist
  'call after loading a picture
  'This routine resets picturebox sizes and Scroll bar values

    With M_MovingPic
        HScrollable = (.Width > m_ContainerPic.Width) 'Test if Scroll is needed
        VScrollable = (.Height > m_ContainerPic.Height)
        If HScrollable Then 'Set TopLeft of picture depending on scroll needed
            .Left = 0 'Needed if previous Picture was smaller than Viewing Area
            XDiff = Abs(.Width - m_ContainerPic.Width)
          Else 'NOT .WIDTH...'HSCROLLABLE = FALSE
            .Left = (m_ContainerPic.Width - .Width) / 2
        End If
        If VScrollable Then
            .Top = 0
            YDiff = Abs(.Height - m_ContainerPic.Height)
          Else 'NOT .HEIGHT...'VSCROLLABLE = FALSE
            .Top = (m_ContainerPic.Height - .Height) / 2
        End If
    End With 'M_MOVINGPIC

    m_ResizingLock = True ' Stops resetting Scroll values from firing Scroll_Change events
    'if not needed then the Max value is set to 0 and the drag button disappears
    With VScrl
        .Max = YDiff
        .LargeChange = IIf(YDiff * m_Scroll_Percent > 0, YDiff * m_Scroll_Percent, 1)
        .Value = 0
        .Enabled = VScrollable
    End With 'VSCRL
    With HScrl
        .Max = XDiff
        .LargeChange = IIf(XDiff * m_Scroll_Percent > 0, XDiff * m_Scroll_Percent, 1)
        .Value = 0
        .Enabled = HScrollable
    End With 'HSCRL
    m_ResizingLock = False
    'set empty value of image for testing if image exists
    InitialPixelValues = PixelSampler

End Sub

Private Sub SetScrolls(pic As PictureBox, ErrMsg As String)
'Copyright 2002 Roger Gilchrist
  'find Scrollbars in same Frame as PictureBox (use to get both own and interlocking scrolls)

  Dim contrl As Control ' used with For Each to find the scrollbars
  Dim ScrlCount As Integer    'escape For Each if both scroolbars are found
  Dim F As Form
  Dim C As Frame
  Dim MyOwnScrolls As Boolean
    If m_ContainerFrame Is Nothing Then 'Safety Trap
        MsgBox "You need to call AssignControls before calling SetScrolls", vbCritical, "ClsFramedScrollingPicture"
        End
    End If
    Set F = pic.Parent
    Set C = pic.Container.Container
    MyOwnScrolls = (C Is m_ContainerFrame) 'test if you are looking for class's own scrollbars or interlocking scrolls

    For Each contrl In F.Controls 'cycle through all controls on form
        If TypeOf contrl Is VScrollBar Then
            If IIf(MyOwnScrolls, VScrl Is Nothing, InterlockVScrl Is Nothing) Then 'if it is a VScroll then
                If contrl.Container Is C Then         'if it is in the Correct Frame
                    If MyOwnScrolls Then
                        Set VScrl = contrl
                      Else 'MYOWNSCROLLS = FALSE
                        Set InterlockVScrl = contrl
                    End If

                    ScrlCount = ScrlCount + 1                        'Count it
                End If
            End If
        End If
        If TypeOf contrl Is HScrollBar Then
            If IIf(MyOwnScrolls, HScrl Is Nothing, InterlockHScrl Is Nothing) Then
                If contrl.Container Is C Then
                    If MyOwnScrolls Then
                        Set HScrl = contrl
                      Else 'MYOWNSCROLLS = FALSE
                        Set InterlockHScrl = contrl
                    End If
                    ScrlCount = ScrlCount + 1
                End If
            End If
        End If
        If ScrlCount = 2 Then ' if two scrollbars have been identifed the stop looking
            Exit For '>---> Next
        End If

    Next contrl
    If ScrlCount <> 2 Then 'Test Requirement 3
        MsgBox ErrMsg & "3", vbCritical, "ClsFramedScrollPicture"
        End
      Else 'NOT SCRLCOUNT...
        If Not MyOwnScrolls Then ' This is used to signal to
            ErrMsg = ""          'AssignInterlockScrolls that you succeeded
        End If

    End If

End Sub

Private Sub VScrl_Change()

  'WithEvents causes this to be called

    If Not m_ResizingLock Then
        M_MovingPic.Top = (-VScrl.Value)
        tmpTop = M_MovingPic.Top
        'move interlocked bar if it exists
        If HasInterlock Then
            InterlockVScrl.Value = VScrl.Value
        End If
    End If

End Sub

Private Sub VScrl_Scroll()

  '  'WithEvents causes this to be called
  '  'This updates image position while you are dragging Scrollbar button
  '  Otherwise the update does not occur until you relase the button

    VScrl_Change

End Sub

':) Ulli's VB Code Formatter V2.13.6 (2/01/2003 8:12:03 PM) 131 + 544 = 675 Lines
