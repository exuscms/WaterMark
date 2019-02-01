VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsOpenSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'--------------------------------------------------------'
' This Code Was Taken From PSC                           '
' Thanks To: Brand-X Software For The Open_File Sub      '
' Minor Edits By Arvinder Sehmi & Creation Of Save_File  '
'Converted to Class and modified Roger Gilchrist Dec 2002'
'email: rojagilkrist@hotmail.com                         '
'Modifications: Subs: NewPicture & Save_Picture added  '
'Modifications" Properties: Loaded, Saving, Filename,    '
'          OpenFilter, SaveFilter, OPenTitle, SaveTitle  '
'          UseGraphicFilters                             '
'TO DO  add Property access to Type OPENFILENAME members '
'       NOTE Some of these are intended only             '
'            for internal use only.                      '
'       Save_As action                                   '
'--------------------------------------------------------'
'
'MEMBERS of Class
'CountChar          --  Private Function. Used to count Chr$(0) to determine max value for FilterIndex
'Filename           --  Read/WriteProperty. Current filename (if any)
'FilenameOnly       --  Read Only Property. Current filename without path. Display purposes only
'FilterFormatter    --  Private Function. Converts Standard '|' filters to the Chr$(0) format used by class
'FilterIndex        --  Read/Write Property. Member of SaveFilter and OpenFilter to display on opening dialog
'FilterIndexOpen    --  Read/Write Property. Member of OpenFilter to display on opening dialog
'FilterIndexSave    --  Read/Write Property. Member of SaveFilter to display on opening dialog
'InitDir            --  Read/Write Property. Initial directory  to display on opening dialog Default = App.Path
'InitDlgs           --  Set up class (call once from Form_Load/Sub Main )
'InitOpenDlg        --  Called from InitDlgs
'InitSaveDlg        --  Called from InitDlgs
'NewPicture       --  Specialized picture Loading wrapper for Open_File
'Loaded             --  Read Only Property. Signal that a file is loaded
'Open_File          --  Show FileOpen dialog
'OpenFilter         --  Set Filter for Open_File    Def="All files|"*.*"
'OpenTitle          --  Set Title for Open_File     Def ="Open File"
'Save_File          --  Show FileSave dialog
'Save_Picture       --  Specialized picture Saveing wrapper for Save_File
'SaveFilter         --  Set Filter for Save_File (may be different to OpenFilter ie you can open more image formats than you can save)
'SaveTitle          --  Set Title for Save_File     Def = "Save File"
'Saving             --  Read Only Property. Signal that a file is being saved. (Prevent close while save is going on)
'UseDocumentFilters --  Write Only Property. Set the OpenFilter and SaveFilter for loading document filetypes
'UseGraphicFilters  --  Write Only Property. Set the OpenFilter and SaveFilter for NewPicture and Save_Picture usage
'
'Declare Api Calls
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Declare Types
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'Declare Variables
Private SaveFileDialog As OPENFILENAME
Private OpenFileDialog As OPENFILENAME
Private m_Loaded As Boolean
Private m_Saving As Boolean
Private m_Filename As String
Private M_SaveFilter As String
Private M_OpenFilter As String
Private m_OpenTitle As String
Private m_SaveTitle As String

Private Sub Class_Initialize()

    OpenFilter = "All Files|*.*|" & _
                 "Graphic Files|*.bmp;*.rle;*.jpg;*.pcx;*.gif;*.ico;*.cur|" & _
                 "Document Files|*.txt;*.doc;*.rtf;*.ini;*.log;*.bat|" & _
                 "Visual Basic Files|*.bas;*.frm;*.vbp;*.vpg*.cls;*.frx;*.vbw;*.dca;*.dsr;*.res"
    'you can add anything else you want as a group here just follow the required format
    '"GroupName|*.ext;*.ext" note ';' separator in file types
    'See UseDocumentFilters and UseGraphicFilters for another way to layout the filter strings
    SaveFilter = M_OpenFilter
    m_OpenTitle = "Open File"
    m_SaveTitle = "Save File"
    FilterIndex = 1
    InitDir = App.Path

End Sub

Private Function CountChar(AStr As String, Test As String) As Long

  'count Test occurring in AStr
  'Uses Split so VB6 specific
  'you can write a Do Loop using Instr to replace this

    CountChar = UBound(Split(AStr, Test))

End Function

Public Property Let Filename(Fname As String)

  'set current filename

    m_Filename = Fname

End Property

Public Property Get Filename() As String

  'retrun current filename if any (Display purposes only)

    Filename = m_Filename

End Property

Public Property Get FilenameOnly() As String

  'retrun current filename if any (Display purposes only)

    FilenameOnly = StrReverse(m_Filename)
    FilenameOnly = StrReverse(Left$(FilenameOnly, InStr(FilenameOnly, "\") - 1))

End Property

Private Function FilterFormatter(StandardFilter As String) As String

  'This allows you to use the standard Filter format using '|' and
  'converts it to the class's required Chr$(0) format

    FilterFormatter = Replace(StandardFilter, "|", Chr$(0))

End Function

Public Property Get FilterIndex() As Long

    FilterIndex = OpenFileDialog.nFilterIndex

End Property

Public Property Let FilterIndex(ByVal Value As Long)

  ' set Open and Save filterIndex to same value
  'Default= 1

    If Value > 0 And Value <= CountChar(OpenFilter, Chr$(0)) Then
        OpenFileDialog.nFilterIndex = Value
    End If
    If Value > 0 And Value <= CountChar(SaveFilter, Chr$(0)) Then
        SaveFileDialog.nFilterIndex = Value
    End If

End Property

Public Property Let FilterIndexOpen(ByVal Value As Long)

  'Allows you to set Open FilterIndex differently to Save FilterIndex

    If Value > 0 And Value <= CountChar(SaveFilter, Chr$(0)) Then
        OpenFileDialog.nFilterIndex = Value
    End If

End Property

Public Property Get FilterIndexOpen() As Long

    FilterIndex = OpenFileDialog.nFilterIndex

End Property

Public Property Get FilterIndexSave() As Long

    FilterIndex = SaveFileDialog.nFilterIndex

End Property

Public Property Let FilterIndexSave(ByVal Value As Long)

  'Allows you to set Save FilterIndex differently to Open FilterIndex

    If Value > 0 And Value <= CountChar(SaveFilter, Chr$(0)) Then
        SaveFileDialog.nFilterIndex = Value
    End If

End Property

Public Property Get InitDir() As String

    InitDir = SaveFileDialog.lpstrInitialDir

End Property

Public Property Let InitDir(ByVal Value As String)

  'Default = App.Path

    SaveFileDialog.lpstrInitialDir = Value

End Property

Public Sub InitDlgs()

    Call InitSaveDlg
    Call InitOpenDlg

End Sub

Private Sub InitOpenDlg()

    With OpenFileDialog
        .lStructSize = Len(OpenFileDialog)
        .hwndOwner = 0&
        .hInstance = App.hInstance
        .lpstrFilter = M_OpenFilter
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = App.Path
        .lpstrTitle = m_OpenTitle
        .flags = 0
    End With 'OPENFILEDIALOG

End Sub

Private Sub InitSaveDlg()

    With SaveFileDialog
        .lStructSize = Len(SaveFileDialog)
        .hwndOwner = 0&
        .hInstance = App.hInstance
        .lpstrFilter = M_SaveFilter
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = App.Path
        .lpstrTitle = m_SaveTitle
        .flags = 0
    End With 'SAVEFILEDIALOG

End Sub

Public Sub Load_Picture(F As Form, P As PictureBox)

    Open_File F.hwnd 'show the open file dlg
    If Loaded Then
        P.Picture = LoadPicture(m_Filename) ' load the file
    End If

End Sub

Public Property Get Loaded() As Boolean

    Loaded = m_Loaded

End Property

Public Function Open_File(hwnd As Long) As String

    If (GetOpenFileName(OpenFileDialog)) Then
        Open_File = Trim$(OpenFileDialog.lpstrFile)
        m_Loaded = True
        m_Filename = Open_File
      Else '(RV&) = FALSE'(GETOPENFILENAME(OPENFILEDIALOG)) = FALSE
        Open_File = ""
        m_Loaded = False
        m_Filename = ""
    End If

End Function

Public Property Get OpenFilter() As String

    OpenFilter = M_OpenFilter

End Property

Public Property Let OpenFilter(ByVal OStr As String)

    M_OpenFilter = FilterFormatter(OStr)
    OpenFileDialog.lpstrFilter = M_OpenFilter

End Property

Public Property Get OpenTitle() As String

    OpenTitle = m_OpenTitle

End Property

Public Property Let OpenTitle(ByVal OStr As String)

    m_OpenTitle = OStr

End Property

Public Function Save_File(hwnd As Long) As String

  Dim tstr As String

    If (GetSaveFileName(SaveFileDialog)) Then
        Save_File = Trim$(SaveFileDialog.lpstrFile)
        tstr = Save_File
      Else '(GETSAVEFILENAME(SAVEFILEDIALOG)) = FALSE
        Save_File = ""
    End If

End Function

Public Sub Save_Picture(F As Form, P As PictureBox)

  Dim StandardCaption As String
  Dim tstr As String

    m_Filename = Save_File(F.hwnd) 'show save dlg
    tstr = m_Filename
    If Trim$(tstr = m_Filename) = "" Then
        tstr = m_Filename
        MsgBox "File Not Saved, Invalid Filename.", vbCritical, "Error"
        Exit Sub ' error in name '>---> Bottom
    End If
    tstr = m_Filename
    m_Saving = True ' start saving
    StandardCaption = F.Caption
    F.Caption = "Saving....."
    P.Picture = P.Image 'set the picture to equal the image
    '-- get rid of any unwanted chars (ie chr13, or 0)
    m_Filename = Replace(m_Filename, Chr$(13), "")
    m_Filename = Replace(m_Filename, Chr$(0), "")
    tstr = m_Filename
    If LCase$(Right$(m_Filename, 4) <> ".bmp") Then
        m_Filename = m_Filename & ".bmp" ' add the bmp on the file
    End If
    tstr = m_Filename
    SavePicture P.Picture, m_Filename  ' save the picture
    m_Saving = False ' no longer saving
    F.Caption = StandardCaption

End Sub

Public Property Get SaveFilter() As String

    SaveFilter = M_SaveFilter

End Property

Public Property Let SaveFilter(ByVal FStr As String)

    M_SaveFilter = FilterFormatter(FStr)
    SaveFileDialog.lpstrFilter = M_SaveFilter

End Property

Public Property Get SaveTitle() As String

    SaveTitle = m_SaveTitle

End Property

Public Property Let SaveTitle(ByVal SStr As String)

    m_SaveTitle = SStr

End Property

Public Property Get Saving() As Boolean

    Saving = m_Saving

End Property

Public Property Let UseDocumentFilters(ByVal UseDFilters As Boolean)

  'You can add any other doc types to this
  'SaveFilter is identical to OpenFilter

    If UseDFilters Then
        OpenFilter = "Document Files|*.txt;*.doc;*.rtf;*.ini;*.log;*.bat|" & _
                     "Text files|*.txt|" & _
                     "Doc files|*.doc|" & _
                     "RTF files|*.rtf|" & _
                     "Bat files|*.bat"
        SaveFilter = OpenFilter
      Else 'USEDFILTERS = FALSE
        Class_Initialize 'Return to Default Filters
    End If

End Property

Public Property Let UseGraphicFilters(ByVal UseGFilters As Boolean)

  'You could build similar routines for other Common filters sets
  'Because this class uses SavePicture which only understands BMP the
  'SaveFilter is different to OpenFilter

    If UseGFilters Then
        OpenFilter = "Image Files|*.bmp;*.jpg;*.pcx;*.gif;*.rle;*.ico;*.cur|" & _
                     "Bitmaps|*.bmp|" & _
                     "Bitmap (RLE)|*.rle|" & _
                     "JPeg files|*.jpg|" & _
                     "Gif Files|*.gif|" & _
                     "Icons|*.ico|" & _
                     "Cursors|*.cur|" & _
                     "PCX files|*.pcx"
        SaveFilter = "Bmp Image File|*.Bmp"
      Else 'USEGFILTERS = FALSE
        Class_Initialize 'Return to Default Filters
    End If

End Property

':) Ulli's VB Code Formatter V2.13.6 (27/12/2002 8:39:02 PM) 80 + 341 = 421 Lines
