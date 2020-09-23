VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl superRTB 
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12270
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   8490
   ScaleWidth      =   12270
   Begin superRTBTest.spellChecker spellChecker1 
      Height          =   4230
      Left            =   6360
      TabIndex        =   38
      Top             =   3960
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7461
      ShowEachWord    =   -1  'True
   End
   Begin VB.CommandButton cmdSPELL 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   295
      Left            =   8520
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Spell Check"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdPreviewIE 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   295
      Left            =   8040
      Picture         =   "superRTB.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Preview In IE"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame frameIMAGE 
      Caption         =   "Insert A Picture"
      Height          =   3975
      Left            =   1080
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   9855
      Begin VB.Frame Frame5 
         Caption         =   "Option Hyperlink (e.g. When image is click this web page will be shown)"
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   5535
         Begin VB.TextBox txtImageLink 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   495
         Left            =   960
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.FileListBox File1 
         Height          =   675
         Left            =   120
         Pattern         =   "*.jpg"
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2535
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Image"
         Height          =   3495
         Left            =   5760
         TabIndex        =   9
         Top             =   360
         Width           =   3975
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   3135
            Left            =   120
            ScaleHeight     =   3105
            ScaleWidth      =   3705
            TabIndex        =   12
            Top             =   240
            Width           =   3735
            Begin VB.Image Image2 
               Height          =   495
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1215
            End
         End
         Begin VB.HScrollBar hs 
            Height          =   255
            LargeChange     =   5
            Left            =   120
            SmallChange     =   10
            TabIndex        =   11
            Top             =   3120
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.VScrollBar vs 
            Height          =   2775
            LargeChange     =   5
            Left            =   3600
            SmallChange     =   10
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.ListBox List1 
         Height          =   2985
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   210
         Left            =   9480
         Picture         =   "superRTB.ctx":03B9
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.Frame frameHyper 
      Caption         =   "Hyperlink"
      Height          =   2175
      Left            =   840
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame Frame3 
         Caption         =   "Visible Link"
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5295
         Begin VB.TextBox txtVisibleLink 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   5055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Link (e.g. http://www.here.com   or  mailto:you@here.com)"
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   5295
         Begin VB.TextBox txtHyperLink 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   5055
         End
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   735
      End
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   1335
      Left            =   5400
      TabIndex        =   32
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
      ExtentX         =   1931
      ExtentY         =   2355
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton cmdITALICOFF 
      Caption         =   "- I"
      Height          =   295
      Left            =   3000
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Turn Italic Off"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdITALICON 
      Caption         =   "+ I"
      Height          =   295
      Left            =   2640
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Turn Italic On"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdPreview 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   295
      Left            =   7680
      Picture         =   "superRTB.ctx":06C4
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Preview On/Off"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdULINEOFF 
      Caption         =   "- U"
      Height          =   295
      Left            =   2280
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Turn Underline Off"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdULINEON 
      Caption         =   "+ U"
      Height          =   295
      Left            =   1920
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Turn Underline On"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdBOLDOFF 
      Caption         =   "- B"
      Height          =   295
      Left            =   1560
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Turn Bold Off"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdBOLDON 
      Caption         =   "+ B"
      Height          =   295
      Left            =   1200
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Turn Bold On"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdCOLOR 
      Caption         =   "Color"
      Height          =   295
      Left            =   600
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Change Font Color"
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cboFONTS 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "cboFONTS"
      ToolTipText     =   "Change Font"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox cboFONTSIZE 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5040
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "cboFONTSIZE"
      ToolTipText     =   "Change Font Size"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdLEFT 
      Height          =   295
      Left            =   5640
      Picture         =   "superRTB.ctx":0A58
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Align Text Left"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdCENTER 
      Height          =   295
      Left            =   6000
      Picture         =   "superRTB.ctx":0DBF
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Align Text Center"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdRIGHT 
      Height          =   295
      Left            =   6360
      Picture         =   "superRTB.ctx":1126
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Align Text Right"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdHYPER 
      Height          =   295
      Left            =   6840
      Picture         =   "superRTB.ctx":148D
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Insert Hyperlink"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdImage 
      Height          =   295
      Left            =   7200
      Picture         =   "superRTB.ctx":186C
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Insert Image"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   2535
      Left            =   120
      TabIndex        =   26
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4471
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   3
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"superRTB.ctx":1C0A
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   8520
      X2              =   10440
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   8640
      X2              =   10560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   9000
      X2              =   10920
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   8640
      X2              =   10560
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "superRTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const imgSPEC = "*.jpg;*.gif;*.bmp"

Public Enum uBackStyles
    Transparent = 0
    Opaque = 1
End Enum

Public Enum uPreviewStyles
    previewFull = 0
    previewHalf = 1
End Enum

Private Const rtfHEAD1 = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033"
Private Const rtfHEAD2 = "\viewkind4\uc1\pard"

Private Type colorType
    Red As Integer
    Blue As Integer
    Green As Integer
End Type


Dim htmPageName As String
Dim rtfHTML As String
Dim htmlImageWidth As Long
Dim htmlImageHeight As Long
Dim fNAME As String
Dim rtbSelstart As Long
Dim MinWidth As Long
Dim DoNotResize As Boolean
Dim PreviewValue As Integer

Dim uAutoGenHTML As Boolean
Dim uPreviewMode As uPreviewStyles
Dim uShowColor As Boolean
Dim uShowItalic As Boolean
Dim uShowBold As Boolean
Dim uShowUline As Boolean
Dim uShowFont As Boolean
Dim uShowAlignments As Boolean
Dim uShowHyperlink As Boolean
Dim uShowPreview As Boolean
Dim uShowBorder As Boolean
Dim uBackColor As Long
Dim uBorderColor As Long
Dim uShowImage As Boolean
Dim uMaxImageWidth As Long
Dim uBackStyle As Integer
Dim uEnabled As Boolean
Dim uShowSpell As Boolean
Dim uLoadAllFonts As Boolean

Public Event Changed()

Private Sub Shell_OpenFile(hWnd As Long, ByVal file As String)
    On Local Error Resume Next
    
    Dim lRet As Long
    Const SW_SHOWNORMAL = 1
    
    lRet = ShellExecute(hWnd, vbNullString, file, vbNullString, App.Path, SW_SHOWNORMAL)

End Sub

Public Sub Generate_HTML()
    On Local Error Resume Next
    Call RTF_To_HTML
End Sub

Public Sub Focus_On_Word(WordPos As Long, FocusWord As String, Optional NewWord As String = "")
    Dim i As Long
    Dim a As Long
    Dim C As Integer
    
    a = WordPos
    
focusAGAIN:
    If a > 0 Then
        rtb.SelStart = a - 1
    Else
        rtb.SelStart = 0
    End If
    rtb.SelLength = Len(FocusWord)
    
    If rtb.SelText <> FocusWord Then
        ' more than likely the word has changed
        ' increment start pos by 1 upto 15 time
        If C <= 15 Then
            a = a + 1
            C = C + 1
            GoTo focusAGAIN
        End If
    End If
    
    If Len(NewWord) <> 0 Then
        rtb.SelText = NewWord
    End If
    
End Sub

Private Sub ShowButtons()
    On Local Error Resume Next
    
    Dim l As Integer
    
    l = 120
   
    If uShowColor = True Then
        cmdCOLOR.Visible = True
        cmdCOLOR.Left = l
        l = l + cmdCOLOR.Width
    Else
        cmdCOLOR.Visible = False
    End If
    
    
    If uShowBold Then
        cmdBOLDON.Left = l
        cmdBOLDON.Visible = True
        l = l + cmdBOLDON.Width
        cmdBOLDOFF.Left = l
        cmdBOLDOFF.Visible = True
        l = l + cmdBOLDOFF.Width
    Else
        cmdBOLDON.Visible = False
        cmdBOLDOFF.Visible = False
    End If
    
    If uShowUline Then
        cmdULINEON.Left = l
        cmdULINEON.Visible = True
        l = l + cmdULINEON.Width
        cmdULINEOFF.Left = l
        cmdULINEOFF.Visible = True
        l = l + cmdULINEOFF.Width
    Else
        cmdULINEON.Visible = False
        cmdULINEOFF.Visible = False
    End If
    
    If uShowItalic Then
        cmdITALICON.Left = l
        cmdITALICON.Visible = True
        l = l + cmdITALICON.Width
        cmdITALICOFF.Left = l
        cmdITALICOFF.Visible = True
        l = l + cmdITALICOFF.Width
    Else
        cmdITALICON.Visible = False
        cmdITALICOFF.Visible = False
    End If
    
    If uShowFont Then
        If l <> 120 Then l = l + 100
        cboFONTS.Left = l
        cboFONTS.Visible = True
        l = l + cboFONTS.Width + 100
        cboFONTSIZE.Left = l
        l = l + cboFONTSIZE.Width
        cboFONTSIZE.Visible = True
    Else
        cboFONTS.Visible = False
        cboFONTSIZE.Visible = False
    End If
    
    If uShowAlignments Then
        If l <> 120 Then l = l + 100
        cmdLEFT.Left = l
        l = l + cmdLEFT.Width
        cmdCENTER.Left = l
        l = l + cmdCENTER.Width
        cmdRIGHT.Left = l
        l = l + cmdRIGHT.Width
        cmdLEFT.Visible = True
        cmdCENTER.Visible = True
        cmdRIGHT.Visible = True
    Else
        cmdLEFT.Visible = False
        cmdCENTER.Visible = False
        cmdRIGHT.Visible = False
    End If
    
    If uShowHyperlink Then
        If l <> 120 Then l = l + 100
        cmdHYPER.Left = l
        l = l + cmdHYPER.Width
        cmdHYPER.Visible = True
    Else
        cmdHYPER.Visible = False
    End If
    
    If uShowImage Then
        cmdImage.Left = l
        l = l + cmdImage.Width
        cmdImage.Visible = True
    Else
        cmdImage.Visible = False
    End If
    
    If uShowPreview Then
        If l <> 120 Then l = l + 100
        cmdPreview.Left = l
        l = l + cmdPreview.Width
        cmdPreviewIE.Left = l
        l = l + cmdPreviewIE.Width
        cmdPreviewIE.Visible = True
        cmdPreview.Visible = True
    Else
        cmdPreview.Visible = False
        cmdPreviewIE.Visible = False
    End If
    
    
    If uShowSpell Then
        If l <> 120 Then l = l + 100
        cmdSPELL.Left = l
        l = l + cmdSPELL.Width
        cmdSPELL.Visible = True
    Else
        cmdSPELL.Visible = False
    End If
    
    MinWidth = l + 120
    
    If uShowImage Then
        MinWidth = 10095
    ElseIf uShowHyperlink Then
        If MinWidth < 5775 Then MinWidth = 5775
    End If
    
    DoNotResize = True
    If UserControl.Width < MinWidth Then UserControl.Width = MinWidth
    If uShowImage And UserControl.Height < 4620 Then UserControl.Height = 4620
    DoNotResize = False
    
    
    If uShowSpell = False And uShowItalic = False And uShowImage = False And uShowColor = False And uShowBold = False And uShowUline = False And uShowFont = False And uShowAlignments = False And uShowHyperlink = False And uShowPreview = False Then
        ' if no buttons visible then make text area larger
        rtb.Top = 120
        rtb.Height = UserControl.Height - 240
    Else
        rtb.Top = 480
        rtb.Height = UserControl.Height - 600
    End If
    Call ShowWebPage
    Call Draw_Borders
    
End Sub

Property Get AutoGenerateHTML() As Boolean
    AutoGenerateHTML = uAutoGenHTML
End Property

Property Let AutoGenerateHTML(uGen As Boolean)
    uAutoGenHTML = uGen
    If uAutoGenHTML Then Call RTF_To_HTML
End Property

Property Get PreviewVisible() As Boolean
    PreviewVisible = web.Visible
End Property

Property Let Enabled(nEnabled As Boolean)
    uEnabled = nEnabled
    Call ED_Control(uEnabled)
End Property

Property Get Enabled() As Boolean
    Enabled = uEnabled
End Property

Property Let PreviewMode(nPreview As uPreviewStyles)
    uPreviewMode = nPreview
    If PreviewValue = 1 Then Call ShowWebPage
End Property

Property Get PreviewMode() As uPreviewStyles
    PreviewMode = uPreviewMode
End Property

Property Get LoadAllFonts() As Boolean
    LoadAllFonts = uLoadAllFonts
End Property

Property Let LoadAllFonts(nFonts As Boolean)
    uLoadAllFonts = nFonts
    Call LoadFonts
End Property

Property Let BackStyle(nBackStyle As uBackStyles)
    uBackStyle = nBackStyle
    UserControl.BackStyle = nBackStyle
    Call Draw_Borders
End Property

Property Get BackStyle() As uBackStyles
    BackStyle = uBackStyle
    Call Draw_Borders
End Property

Property Let ShowItalic(nItalic As Boolean)
    uShowItalic = nItalic
    Call ShowButtons
End Property

Property Get ShowItalic() As Boolean
    ShowItalic = uShowItalic
End Property

Property Let ShowInsertImage(nImage As Boolean)
    uShowImage = nImage
    Call ShowButtons
End Property

Property Get ShowInsertImage() As Boolean
    ShowInsertImage = uShowImage
End Property

Property Let ShowSpellChecker(nSpell As Boolean)
    uShowSpell = nSpell
    Call ShowButtons
End Property

Property Get ShowSpellChecker() As Boolean
    ShowSpellChecker = uShowSpell
End Property

Property Let MaxImageWidth(nWidth As Long)
    uMaxImageWidth = nWidth
End Property

Property Get MaxImageWidth() As Long
    MaxImageWidth = uMaxImageWidth
End Property

Property Let BorderColor(nColor As OLE_COLOR)
    uBorderColor = nColor
    Call Draw_Borders
End Property

Property Get BorderColor() As OLE_COLOR
    BorderColor = uBorderColor
End Property

Property Let BackColor(nColor As OLE_COLOR)
    uBackColor = nColor
    UserControl.BackColor = nColor
End Property

Property Get BackColor() As OLE_COLOR
    BackColor = uBackColor
    Call Draw_Borders
End Property

Property Let ShowBorder(nBorder As Boolean)
    uShowBorder = nBorder
    Call Draw_Borders
End Property

Property Get ShowBorder() As Boolean
    ShowBorder = uShowBorder
End Property

Property Let ShowPreview(nPreview As Boolean)
    uShowPreview = nPreview
    Call ShowButtons
End Property

Property Get ShowPreview() As Boolean
    ShowPreview = uShowPreview
End Property
Property Let ShowHyperlink(nLink As Boolean)
    uShowHyperlink = nLink
    Call ShowButtons
End Property

Property Get ShowHyperlink() As Boolean
    ShowHyperlink = uShowHyperlink
End Property
Property Let ShowAlignments(nAlign As Boolean)
    uShowAlignments = nAlign
    Call ShowButtons
End Property

Property Get ShowAlignments() As Boolean
    ShowAlignments = uShowAlignments
End Property

Property Let ShowFont(nFont As Boolean)
    uShowFont = nFont
    Call ShowButtons
End Property

Property Get ShowFont() As Boolean
    ShowFont = uShowFont
End Property

Property Let ShowUline(nUline As Boolean)
    uShowUline = nUline
    Call ShowButtons
End Property

Property Get ShowUline() As Boolean
    ShowUline = uShowUline
End Property

Property Let ShowColor(nColor As Boolean)
    uShowColor = nColor
    Call ShowButtons
End Property

Property Get ShowColor() As Boolean
    ShowColor = uShowColor
End Property

Property Let ShowBold(nBold As Boolean)
    uShowBold = nBold
    Call ShowButtons
End Property

Property Get ShowBold() As Boolean
    ShowBold = uShowBold
End Property

Private Function Replace_Text(searchSTRING As String, searchFOR As String, replaceWITH As String) As String
    Dim newSTRING As String
    Dim i As Long
    Dim lPART As String
    Dim rPART As String
    
    On Local Error Resume Next
    
    newSTRING = searchSTRING
    i = InStr(1, newSTRING, searchFOR, vbBinaryCompare)
    While i <> 0
        lPART = Left$(newSTRING, i - 1)
        rPART = Right$(newSTRING, Len(newSTRING) - ((i - 1) + Len(searchFOR)))
        newSTRING = lPART & replaceWITH & rPART
        i = InStr(i + Len(replaceWITH), newSTRING, searchFOR, vbBinaryCompare)
    Wend
    Replace_Text = newSTRING
    
End Function

Private Sub HTML_To_RTF()
    Dim tRTF As String
    Dim i As Long
    Dim a As Long
    Dim spanFONT As String
    Dim spanSIZE As Integer
    Dim spanRTF As String
    Dim colorTABLE As String
    Dim fontCOLOR1 As String
    Dim fontCOLOR2 As String
    Dim colorID As Long
    Dim cR As Long
    Dim cG As Long
    Dim cB As Long
    Dim fontTABLE As String
    Dim fontITEMS() As String
    Dim fontCNT As Long
    Dim tFONT As String
    Dim C As Long
    Dim newALIGN As String
    
    On Local Error Resume Next
    
    tRTF = rtfHTML
    
    ReDim fontITEMS(0 To 0)
    
    fontTABLE = "{\fonttbl"
    
    i = InStr(1, tRTF, "font-family:", vbTextCompare)
    While i <> 0
        a = InStr(i + 13, tRTF, Chr$(34), vbTextCompare)
        tFONT = Mid$(tRTF, i + 13, a - (i + 13))
        For C = 1 To UBound(fontITEMS)
            If Right$(fontITEMS(C), Len(tFONT)) = tFONT Then Exit For
        Next C
        If C > UBound(fontITEMS) Then
            fontTABLE = fontTABLE & "{\f" & fontCNT & "\fnil\fcharset0 " & tFONT & ";}"
            fontCNT = fontCNT + 1
            ReDim Preserve fontITEMS(0 To fontCNT)
            fontITEMS(fontCNT) = "\f" & (fontCNT - 1) & " " & tFONT
        End If
    
        i = InStr(i + 1, tRTF, "font-family:", vbTextCompare)
    Wend
    fontTABLE = fontTABLE & "}"
    
'{\fonttbl{\f0\fnil\fcharset0 Arial;}{\f1\fnil\fcharset0 Bookman Old Style;}{\f2\fnil\fcharset0 MS Sans Serif;}}
'<span style='font-size:8.0pt;font-family:"Microsoft Sans Serif"'>normaltext<br />

    i = InStr(1, tRTF, "<FONT COLOR=", vbTextCompare)
    While i <> 0
        a = InStr(i + 1, tRTF, ">")
        fontCOLOR1 = Mid$(tRTF, i, (a - i) + 1)
        fontCOLOR2 = Replace_Text(fontCOLOR1, "<FONT COLOR=", "")
        fontCOLOR2 = Replace_Text(fontCOLOR2, Chr$(34), "")
        fontCOLOR2 = Replace_Text(fontCOLOR2, ">", "")
        fontCOLOR2 = Replace_Text(fontCOLOR2, "#", "")

        If Len(colorTABLE) = 0 Then
            colorTABLE = "{\colortbl ;"
            colorID = 1
        Else
            colorID = colorID + 1
        End If
        
        cR = Val("&H" & Left$(fontCOLOR2, 2))
        cG = Val("&H" & Mid$(fontCOLOR2, 3, 2))
        cB = Val("&H" & Right$(fontCOLOR2, 2))
        
        If cR = 0 And cG = 0 And cB = 0 Then
            colorID = colorID - 1
            tRTF = Replace_Text(tRTF, fontCOLOR1, "\cf0")
        Else
            colorTABLE = colorTABLE & "\red" & cR & "\green" & cG & "\blue" & cB & ";"
            tRTF = Replace_Text(tRTF, fontCOLOR1, "\cf" & colorID)
        End If
        i = InStr(1, tRTF, "<FONT COLOR=", vbTextCompare)
    Wend
    
    If Len(colorTABLE) <> 0 Then colorTABLE = colorTABLE & "}" & vbCrLf
    
    'tabs
    tRTF = Replace_Text(tRTF, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ", "\tab ")
    tRTF = Replace_Text(tRTF, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", "\tab ")
    tRTF = Replace_Text(tRTF, "&nbsp;", " ")
    
    tRTF = Replace_Text(tRTF, vbCrLf, "")
    tRTF = Replace_Text(tRTF, "<br>", "\par ")
    tRTF = Replace_Text(tRTF, "<br />", "\par ")
    tRTF = Replace_Text(tRTF, "</I>", "\i0")
    tRTF = Replace_Text(tRTF, "<I>", "\i ")
    tRTF = Replace_Text(tRTF, "</b>", "\b0")
    tRTF = Replace_Text(tRTF, "<b>", "\b ")
    tRTF = Replace_Text(tRTF, "</u>", "\ulnone")
    tRTF = Replace_Text(tRTF, "<u>", "\ul ")
    tRTF = Replace_Text(tRTF, "</span>", "")
    tRTF = Replace_Text(tRTF, "</FONT>", "\cf0")
    tRTF = Replace_Text(tRTF, "&amp;", "&")
    tRTF = Replace_Text(tRTF, "&ndash;", "-")
    tRTF = Replace_Text(tRTF, "&pound;", "\'a3")
    tRTF = Replace_Text(tRTF, "£", "\'a3")
    tRTF = Replace_Text(tRTF, "&squo;", "'")
    
    i = InStr(1, tRTF, "<span style", vbTextCompare)
    While i <> 0
        a = InStr(i, tRTF, ">")
        spanFONT = Mid$(tRTF, i, (a - i) + 1)
        
        a = InStr(1, spanFONT, ":")
        spanSIZE = Val(Mid$(spanFONT, a + 1, 3))
        spanRTF = "\fs" & spanSIZE * 2 & " "
        
        a = InStr(1, tRTF, "font-family:", vbTextCompare)
        C = InStr(a + 13, tRTF, Chr$(34), vbTextCompare)
        tFONT = Mid$(tRTF, a + 13, C - (a + 13))
        For a = 1 To UBound(fontITEMS)
            If Right$(fontITEMS(a), Len(tFONT)) = tFONT Then Exit For
        Next a
        tFONT = "\f" & (a - 1)
        spanRTF = tFONT & spanRTF
        
        
        tRTF = Replace_Text(tRTF, spanFONT, spanRTF)
        
    
        i = InStr(i, tRTF, "<span style", vbTextCompare)
    
    Wend

    ' do alignments
    
    
    tRTF = Replace_Text(tRTF, "<P ALIGN=LEFT></P>", "")
    tRTF = Replace_Text(tRTF, "<P ALIGN=RIGHT></P>", "")
    tRTF = Replace_Text(tRTF, "<P ALIGN=CENTER></P>", "")
    
    tRTF = Replace_Text(tRTF, "<P ALIGN=CENTER> ", " \pard\qc ")
    tRTF = Replace_Text(tRTF, "<P ALIGN=CENTER>", " \pard\qc ")
    tRTF = Replace_Text(tRTF, "<P ALIGN=RIGHT> ", " \pard\qr ")
    tRTF = Replace_Text(tRTF, "<P ALIGN=RIGHT>", " \pard\qr ")
    
    tRTF = Replace_Text(tRTF, "<P ALIGN=LEFT> ", " \pard ")
    tRTF = Replace_Text(tRTF, "<P ALIGN=LEFT>", " \pard ")
    
    tRTF = Replace_Text(tRTF, "</P>", "\par ")
    
    tRTF = Replace_Text(tRTF, "\pard \par \pard", "\par\pard")
    tRTF = Replace_Text(tRTF, "\par \pard", "\par\pard")
    
'    tRTF = rtfHEAD1 & fontTABLE & colorTABLE & rtfHEAD2 & "\fi-795\li795\tx795 " & tRTF & "\par }"
    tRTF = rtfHEAD1 & fontTABLE & colorTABLE & rtfHEAD2 & "\fi-795\tx795 " & tRTF & "\par }"
    rtb.TextRTF = tRTF

'<span style='font-size:10.0pt;font-family:"Arial"'> left<br /><P ALIGN=LEFT></P><P ALIGN=CENTER> center</P><P ALIGN=LEFT></P><P ALIGN=RIGHT> right</P><P ALIGN=LEFT></span><span style='font-size:8.0pt;font-family:"MS Sans Serif"'> </P>
End Sub

Private Sub RTF_To_HTML()
    Dim tRTF As String
    Dim hHTML As String
    Dim i As Long
    Dim a As Long
    Dim spanSIZE As String
    Dim firstFONT As Boolean
    Dim rtfCOLORS() As colorType
    Dim colorSTRIP As String
    Dim B As Long
    Dim colorITEM As String
    Dim hCOLOR As String
    Dim rtfFONTS() As String
    Dim tFONT As String
    Dim fontSIZE As String
    Dim partA As String
    Dim partB As String
    Dim newRTF As String
    Dim firstALIGN As Boolean
    Dim newALIGN As String
    Dim langCODE As String
    
    On Local Error Resume Next
    
    ' only copes with one font at the moment
    tRTF = rtb.TextRTF
    ReDim rtfFONTS(0 To 0)
       
    ' remove head
    tRTF = Replace_Text(tRTF, rtfHEAD1, "")
    tRTF = Replace_Text(tRTF, rtfHEAD2, "")
    tRTF = Replace_Text(tRTF, "{\rtf1\ansi\deff0", "")
    tRTF = Replace_Text(tRTF, "\sb100", "")
    tRTF = Replace_Text(tRTF, "\sa100", "")
    tRTF = Replace_Text(tRTF, "\pard\fi720\qj", "\tab")
    tRTF = Replace_Text(tRTF, "\pard\qj\tab", "\tab")
    tRTF = Replace_Text(tRTF, "\qj", "")
    tRTF = Replace_Text(tRTF, "\sl280", "")
    tRTF = Replace_Text(tRTF, "\slmult0", "")
    tRTF = Replace_Text(tRTF, "\expndtw-25", "")
    tRTF = Replace_Text(tRTF, "\expndtw0", "")
    tRTF = Replace_Text(tRTF, "\sl280", "")
    
    tRTF = Replace_Text(tRTF, "\tx795", "")
    tRTF = Replace_Text(tRTF, "\tx0", "")
    tRTF = Replace_Text(tRTF, "\fi-795", "")
    tRTF = Replace_Text(tRTF, "\li795", "")
    tRTF = Replace_Text(tRTF, "\fi720", "")
    
    tRTF = Replace_Text(tRTF, "\ldblquote", Chr$(34))
    tRTF = Replace_Text(tRTF, "\ldblquote ", Chr$(34))
    tRTF = Replace_Text(tRTF, "\rdblquote", Chr$(34))
    tRTF = Replace_Text(tRTF, "\enddash", "-")
    tRTF = Replace_Text(tRTF, "\rquote ", "'")
    tRTF = Replace_Text(tRTF, "\rquote", "'")
    tRTF = Replace_Text(tRTF, "\lquote ", "'")
    tRTF = Replace_Text(tRTF, "\lquote", "'")
    
    For i = 0 To 500
        tRTF = Replace_Text(tRTF, "\fs" & i & " ", "\fs" & i)
    Next i
    
    ' highlights are not allow
    For i = 0 To 99
        B = InStr(1, tRTF, "\highlight", vbTextCompare)
        If B = 0 Then Exit For
        tRTF = Replace_Text(tRTF, "\highlight" & i, "")
    Next i
    
    ' replace and \lang
    i = InStr(1, tRTF, "\lang", vbTextCompare)
    While i <> 0
        langCODE = Mid$(tRTF, i + 5, 4)
        If IsNumeric(langCODE) Then
            tRTF = Replace_Text(tRTF, "\lang" & langCODE, "")
        End If
        i = InStr(i + 1, tRTF, "\lang", vbTextCompare)
    Wend
    
    
    'tRTF = Replace_Text(tRTF, "\lang2057", "")
    'tRTF = Replace_Text(tRTF, "\lang1033", "")
    
    tRTF = Replace_Text(tRTF, vbCrLf, "")
    ReDim rtfCOLORS(0 To 0)
    
    i = InStr(1, tRTF, "{\colortbl ;", vbTextCompare)
    If i > 0 Then
        a = InStr(i + 1, tRTF, "}", vbTextCompare)
        colorSTRIP = Mid$(tRTF, i, (a - (i - 1)))
        tRTF = Replace_Text(tRTF, colorSTRIP, "")
        
        colorSTRIP = Replace_Text(colorSTRIP, "{\colortbl ;", "")
        colorSTRIP = Replace_Text(colorSTRIP, "}", "")
        
        i = InStr(1, colorSTRIP, ";", vbTextCompare)
        a = 0
        While i <> 0
            a = a + 1
            ReDim Preserve rtfCOLORS(0 To a)
            
            colorITEM = Left$(colorSTRIP, i)
            colorSTRIP = Replace_Text(colorSTRIP, colorITEM, "")
            
            colorITEM = Right$(colorITEM, Len(colorITEM) - 1)
            For B = 1 To 3
                If Left$(colorITEM, 3) = "red" Then
                    colorITEM = Right$(colorITEM, Len(colorITEM) - 3)
                    i = InStr(1, colorITEM, "\")
                    rtfCOLORS(a).Red = Val(Left$(colorITEM, i - 1))
                ElseIf Left$(colorITEM, 5) = "green" Then
                    colorITEM = Right$(colorITEM, Len(colorITEM) - 5)
                    i = InStr(1, colorITEM, "\")
                    rtfCOLORS(a).Green = Val(Left$(colorITEM, i - 1))
                Else
                    colorITEM = Right$(colorITEM, Len(colorITEM) - 4)
                    i = InStr(1, colorITEM, ";")
                    rtfCOLORS(a).Blue = Val(Left$(colorITEM, i - 1))
                End If
                colorITEM = Right$(colorITEM, Len(colorITEM) - i)
            Next B
        
            i = InStr(1, colorSTRIP, ";", vbTextCompare)
        Wend
        
'{\colortbl ;\red255\green0\blue0;}
    End If
    
    ' do font table
    i = InStr(1, tRTF, "{\fonttbl", vbTextCompare)
    If i > 0 Then
        a = InStr(i + 1, tRTF, ";}}", vbTextCompare)
        tFONT = Mid$(tRTF, i, (a - i) + 3)
        ' will have a ;}} closing
        tRTF = Replace_Text(tRTF, tFONT, "")
        tFONT = Replace_Text(tFONT, "{\fonttbl", "")
        tFONT = Left$(tFONT, Len(tFONT) - 1)
    
        a = 0
        i = InStr(1, tFONT, "}", vbTextCompare)
        While i > 0
            a = a + 1
            ReDim Preserve rtfFONTS(0 To a)
            rtfFONTS(a) = Left$(tFONT, i)
            B = InStr(1, rtfFONTS(a), " ", vbTextCompare)
            rtfFONTS(a) = Left$(rtfFONTS(a), 4) & " " & Right$(rtfFONTS(a), Len(rtfFONTS(a)) - B)
            rtfFONTS(a) = Replace_Text(rtfFONTS(a), "\fnil\fcharset0", "")
            rtfFONTS(a) = Replace_Text(rtfFONTS(a), "{\", "")
            rtfFONTS(a) = Replace_Text(rtfFONTS(a), ";}", "")
            
            tFONT = Right$(tFONT, Len(tFONT) - i)
            i = InStr(1, tFONT, "}", vbTextCompare)
        Wend
    End If
    
    'tabs
    tRTF = Replace_Text(tRTF, "\tab", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
    tRTF = Replace_Text(tRTF, "\tab ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
  
  
  
   ' at the end is always a  }
    tRTF = Left$(tRTF, Len(tRTF) - 1)
    
    tFONT = "\f0"
    newRTF = ""
    fontSIZE = "\fs20"
    For i = 1 To Len(tRTF)
        If Mid$(tRTF, i, 2) = "\f" Then
            ' could be \f0 or \fs
            If Mid$(tRTF, i, 3) = "\fs" Then
                newRTF = newRTF & tFONT & "\fs"
                fontSIZE = Mid$(tRTF, i, 5)
                i = i + 2
                GoTo skipNEXT
            Else
                ' font selection
                If Mid$(tRTF, i - 5, 3) <> "\fs" And Mid$(tRTF, i + 3, 3) <> "\fs" Then
                    tFONT = Mid$(tRTF, i, 3)
                    newRTF = newRTF & tFONT & fontSIZE
                    i = i + 2
                    GoTo skipNEXT
                Else
                    tFONT = Mid$(tRTF, i, 3)
                    i = i + 2
                    GoTo skipNEXT
                End If
                
            End If
        Else
            newRTF = newRTF & Mid$(tRTF, i, 1)
        End If
    
skipNEXT:
    Next i
    
    tRTF = newRTF
    
    firstFONT = True
    i = InStr(1, tRTF, "\f")
    While i <> 0
        ' get front name
        tFONT = Trim$(Mid$(tRTF, i + 1, 2))
        For B = 1 To UBound(rtfFONTS)
            If tFONT = Left$(rtfFONTS(B), 2) Then
                tFONT = Right$(rtfFONTS(B), Len(rtfFONTS(B)) - 3)
                Exit For
            End If
        Next B
    
    
        a = Mid$(tRTF, i + 6, 2)
        spanSIZE = "<span style='font-size:" & a \ 2 & ".0pt;font-family:" & Chr$(34) & tFONT & Chr$(34) & "'>"
        If Not firstFONT Then
            spanSIZE = "</span>" & spanSIZE
        End If

        tRTF = Left$(tRTF, i - 1) & spanSIZE & Right$(tRTF, Len(tRTF) - (i + 7))
        i = i + Len(spanSIZE)
        i = InStr(i, tRTF, "\f")
        firstFONT = False
    Wend
    
    ' lets do colors
'<FONT COLOR="#ffff00">
    If UBound(rtfCOLORS) > 0 Then
        For i = 1 To UBound(rtfCOLORS)
            hCOLOR = RGB_To_HTML_Color(rtfCOLORS(i).Red, rtfCOLORS(i).Green, rtfCOLORS(i).Blue)
            If i = 1 Then
                tRTF = Replace_Text(tRTF, "\cf" & i & " ", "<FONT COLOR=" & Chr$(34) & hCOLOR & Chr$(34) & ">")
                tRTF = Replace_Text(tRTF, "\cf" & i, "<FONT COLOR=" & Chr$(34) & hCOLOR & Chr$(34) & ">")
            Else
                tRTF = Replace_Text(tRTF, "\cf" & i & " ", "</FONT><FONT COLOR=" & Chr$(34) & hCOLOR & Chr$(34) & ">")
                tRTF = Replace_Text(tRTF, "\cf" & i, "</FONT><FONT COLOR=" & Chr$(34) & hCOLOR & Chr$(34) & ">")
            End If
        Next i
        tRTF = Replace_Text(tRTF, "\cf0", "</FONT>")
    End If
        
    ' do alignments
    ' left aligns are encoded as \pard, replace with our own of \ql
    tRTF = Replace_Text(tRTF, "\pard", "\ql")
    i = InStr(1, tRTF, "\q", vbTextCompare)
    firstALIGN = True
    While i <> 0
        If Mid$(tRTF, i + 2, 1) = "r" Or Mid$(tRTF, i + 2, 1) = "c" Or Mid$(tRTF, i + 2, 1) = "l" Then
            ' found a alignmment
            
            newALIGN = ""
            If Not firstALIGN Then newALIGN = "</P>"
            
            If Mid$(tRTF, i + 2, 1) = "r" Then
                newALIGN = newALIGN & "<P ALIGN=RIGHT>"
            ElseIf Mid$(tRTF, i + 2, 1) = "c" Then
                newALIGN = newALIGN & "<P ALIGN=CENTER>"
            ElseIf Mid$(tRTF, i + 2, 1) = "l" Then
                newALIGN = newALIGN & "<P ALIGN=LEFT>"
            End If
            
            tRTF = Left$(tRTF, i - 1) & newALIGN & Right$(tRTF, Len(tRTF) - (i + 2))
            
            If firstALIGN And Len(newALIGN) = 0 Then i = i - 3
            firstALIGN = False
        End If
    
        i = InStr(i + 1, tRTF, "\q", vbTextCompare)
    Wend
    If firstALIGN = False Then
        ' have done an alignment
        tRTF = tRTF & "</P>"
    End If
    
    
    
    tRTF = Replace_Text(tRTF, vbCrLf & "</P>", "</P>")
    tRTF = Replace_Text(tRTF, "</P>" & vbCrLf, "</P>")
    tRTF = Replace_Text(tRTF, "</P></P>", "</P>")
    tRTF = Replace_Text(tRTF, "\par ", "<br />") ' & vbCrLf)
    tRTF = Replace_Text(tRTF, "<br /></P>", "</P>")
    
    tRTF = Replace_Text(tRTF, "\i0", "</I>")
    tRTF = Replace_Text(tRTF, "\i0 ", "</I>")
    tRTF = Replace_Text(tRTF, "\i", "<I>")
    tRTF = Replace_Text(tRTF, "\i ", "<I>")
    
    tRTF = Replace_Text(tRTF, "\b0", "</b>")
    tRTF = Replace_Text(tRTF, "\b0 ", "</b>")
    tRTF = Replace_Text(tRTF, "\b", "<b>")
    tRTF = Replace_Text(tRTF, "\b ", "<b>")
    tRTF = Replace_Text(tRTF, "\ulnone", "</u>")
    tRTF = Replace_Text(tRTF, "\ulnone ", "</u>")
    tRTF = Replace_Text(tRTF, "\ul", "<u>")
    tRTF = Replace_Text(tRTF, "\ul ", "<u>")
    
    tRTF = Replace_Text(tRTF, "\'a3", "&pound;")
    
    tRTF = Replace_Text(tRTF, "\\", "\")
    tRTF = Replace_Text(tRTF, "//", "/")
    
    ' reformat HTTP links
    i = InStr(1, tRTF, "http:", vbTextCompare)
    While i <> 0
        partA = Left$(tRTF, i + 4) & "//"
        a = i + 5
        For B = a To Len(tRTF)
            If Mid$(tRTF, B, 1) <> "/" And Mid$(tRTF, B, 1) <> "\" Then Exit For
        Next B
        partB = Right$(tRTF, Len(tRTF) - (B - 1))
        tRTF = partA & partB
        i = InStr(B, tRTF, "http:", vbTextCompare)
    Wend
    
    
    ' reformat HTTPS links
    i = InStr(1, tRTF, "https:", vbTextCompare)
    While i <> 0
        partA = Left$(tRTF, i + 5) & "//"
        a = i + 6
        For B = a To Len(tRTF)
            If Mid$(tRTF, B, 1) <> "/" And Mid$(tRTF, B, 1) <> "\" Then Exit For
        Next B
        partB = Right$(tRTF, Len(tRTF) - (B - 1))
        tRTF = partA & partB
        i = InStr(B, tRTF, "https:", vbTextCompare)
    Wend
    
    tRTF = Replace_Text(tRTF, "&nb! sp;", "&nbsp;")
    
    rtfHTML = tRTF
    
'\par \fs16 font8
'\par \fs20 font10
'\par \fs24 font12
'\par \fs28 font14
'\par \fs32 font16
'\par \fs17
'\par }

'<span style='font-size:14.0pt;font-family:"Microsoft Sans Serif"'>Size14</span>
'<span style='font-size:16.0pt;font-family:"Microsoft Sans Serif"'>Size16</span>
'<span style='font-size:18.0pt;font-family:"Microsoft Sans Serif"'>Size18</span>
'<span style='font-size:20.0pt;font-family:"Microsoft Sans Serif"'>Size20</span>
'<span style='font-size:22.0pt;font-family:"Microsoft Sans Serif"'>Size22</span>


End Sub

Private Function RGB_To_HTML_Color(iRed As Integer, iGreen As Integer, iBlue As Integer) As String

    Dim hR As String
    Dim hG As String
    Dim hB As String
    
    hR = Right$("0" & Hex$(iRed), 2)
    hG = Right$("0" & Hex$(iGreen), 2)
    hB = Right$("0" & Hex$(iBlue), 2)
    
    RGB_To_HTML_Color = "#" & hR & hG & hB

End Function

Property Get TextHTML() As String
    On Local Error Resume Next
    TextHTML = rtfHTML
End Property

Property Let TextHTML(nHTML As String)
    On Local Error Resume Next
    rtfHTML = nHTML
    Call HTML_To_RTF
End Property

Property Get Text() As String
    On Local Error Resume Next
    Text = rtb.Text
End Property

Property Let Text(nText As String)
    On Local Error Resume Next
    rtb.Text = nText
End Property

Property Get TextRTF() As String
    On Local Error Resume Next
    TextRTF = rtb.TextRTF
End Property

Property Let TextRTF(nText As String)
    On Local Error Resume Next
    rtb.TextRTF = nText
End Property

Private Sub cboFONTS_Click()
    On Local Error Resume Next
    rtb.SelFontName = cboFONTS.Text
    rtb.SetFocus
    RaiseEvent Changed
End Sub

Private Sub cboFONTS_KeyDown(KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    KeyCode = 0
End Sub

Private Sub cboFONTS_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    KeyAscii = 0
End Sub

Private Sub cboFONTSIZE_Click()
    On Local Error Resume Next
    rtb.SelFontSize = cboFONTSIZE.Text
    rtb.SetFocus
    RaiseEvent Changed
End Sub

Private Sub cboFONTSIZE_KeyDown(KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    KeyCode = 0
End Sub

Private Sub cboFONTSIZE_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    KeyAscii = 0
End Sub



Private Sub cmdApply_Click()
    Dim aLINK As String
    
    On Local Error Resume Next
    
    aLINK = "<A HREF=" & Chr$(34) & txtHyperLink.Text & Chr$(34) & ">" & txtVisibleLink & "</A>"
    
    frameHyper.Visible = False
    
    Call ED_Control(True)
    rtb.SelStart = rtbSelstart
    rtb.SelText = aLINK
    Call rtb_Change
    rtb.SetFocus
    
    
End Sub

Private Sub cmdBOLDOFF_Click()
    On Local Error Resume Next
    rtb.SelBold = False
    rtb.SetFocus
    RaiseEvent Changed
End Sub

Private Sub cmdBOLDON_Click()
    On Local Error Resume Next
    rtb.SelBold = True
    rtb.SetFocus
    RaiseEvent Changed
End Sub

Private Sub cmdCancel_Click()
    On Local Error Resume Next
    frameHyper.Visible = False
    
    Call ED_Control(True)
    
    rtb.SetFocus
End Sub

Private Sub cmdCENTER_Click()
    On Local Error Resume Next
    rtb.SelAlignment = 2
    rtb.SetFocus
    RaiseEvent Changed
End Sub

Private Sub cmdCOLOR_Click()
    On Local Error Resume Next
    
    CommonDialog1.ShowColor
    rtb.SelColor = CommonDialog1.Color
    rtb.SetFocus
    RaiseEvent Changed
End Sub

Private Sub cmdHYPER_Click()
    On Local Error Resume Next
    
    rtbSelstart = rtb.SelStart
    
    frameHyper.Move (UserControl.Width - frameHyper.Width) / 2, (UserControl.Height - frameHyper.Height) / 2

    Call ED_Control(False)
    
    txtVisibleLink.Text = ""
    txtHyperLink.Text = ""
    
    frameHyper.Visible = True
    
    txtVisibleLink.SetFocus
End Sub

Private Sub ED_Control(ed As Boolean, Optional OmitPreview As Boolean = False)
    On Local Error Resume Next
    
    cmdImage.Enabled = ed
    If OmitPreview = False Then cmdPreview.Enabled = ed
    cmdITALICON.Enabled = ed
    cmdITALICOFF.Enabled = ed
    cmdCOLOR.Enabled = ed
    cmdBOLDON.Enabled = ed
    cmdBOLDOFF.Enabled = ed
    cmdULINEON.Enabled = ed
    cmdULINEOFF.Enabled = ed
    cboFONTS.Enabled = ed
    cboFONTSIZE.Enabled = ed
    cmdLEFT.Enabled = ed
    cmdCENTER.Enabled = ed
    cmdRIGHT.Enabled = ed
    cmdHYPER.Enabled = ed
    cmdSPELL.Enabled = ed
    cmdPreviewIE.Enabled = ed
    rtb.Enabled = ed

End Sub

Private Sub cmdImage_Click()
    On Local Error Resume Next
    
    rtbSelstart = rtb.SelStart
    frameIMAGE.Move (UserControl.Width - frameIMAGE.Width) / 2, (UserControl.Height - frameIMAGE.Height) / 2

    Call ED_Control(False)
    
    txtImageLink.Text = ""
    
    fNAME = ""
    Call Dir1_Change
    
    frameIMAGE.Visible = True
    
End Sub


Private Sub cmdITALICOFF_Click()
    On Local Error Resume Next
    rtb.SelItalic = False
    rtb.SetFocus
    RaiseEvent Changed
End Sub

Private Sub cmdITALICON_Click()
    On Local Error Resume Next
    rtb.SelItalic = True
    rtb.SetFocus
    RaiseEvent Changed
End Sub

Private Sub cmdLEFT_Click()
    On Local Error Resume Next
    rtb.SelAlignment = 0
    rtb.SetFocus
    RaiseEvent Changed
End Sub

Private Sub cmdPreview_Click()
    On Local Error Resume Next
    If PreviewValue = 0 Then
        PreviewValue = 1
    Else
        PreviewValue = 0
    End If
    Call ShowWebPage
    rtb.SetFocus
End Sub

Private Sub ShowWebPage()
    
    Dim hFN As Integer
    On Local Error Resume Next
    If PreviewValue = 0 Then
        If uPreviewMode = previewFull Then
            Call ED_Control(True, True)
        End If
        rtb.Width = UserControl.Width - 240
        web.Visible = False
    Else
        If uPreviewMode = previewHalf Then
            rtb.Width = (UserControl.Width - 480) / 2
            web.Top = rtb.Top
            web.Left = rtb.Left + rtb.Width + 240
            web.Height = rtb.Height
            web.Width = UserControl.Width - (rtb.Left + rtb.Width + 360)
        Else
            web.Left = rtb.Left
            web.Width = UserControl.Width - 240
            web.Height = rtb.Height
            web.Top = rtb.Top
            
            Call ED_Control(False, True)
        End If
        
        If uAutoGenHTML = False Then Call RTF_To_HTML
        
        web.Visible = True
        
        If Len(htmPageName) = 0 Then
            htmPageName = App.Path & "\" & UserControl.Ambient.DisplayName & ".htm"
        End If
        hFN = FreeFile
        Open htmPageName For Output As #hFN
        Print #hFN, "<html>"
        Print #hFN, "<head>"
        Print #hFN, "</head>"
        Print #hFN, "<body>"
        Print #hFN, rtfHTML
        Print #hFN, "</body>"
        Print #hFN, "</html>"
        Close #hFN

        web.Navigate htmPageName
    End If

End Sub

Private Sub cmdPreviewIE_Click()
    Dim hFN As Integer
    
    On Local Error Resume Next
    
    If Len(htmPageName) = 0 Then
        htmPageName = App.Path & "\" & UserControl.Ambient.DisplayName & ".htm"
    End If
    hFN = FreeFile
    Open htmPageName For Output As #hFN
    Print #hFN, "<html>"
    Print #hFN, "<head>"
    Print #hFN, "</head>"
    Print #hFN, "<body>"
    Print #hFN, rtfHTML
    Print #hFN, "</body>"
    Print #hFN, "</html>"
    Close #hFN


    Call Shell_OpenFile(UserControl.Parent.hWnd, htmPageName)
'    Call Shell_OpenFile(Form1.hWnd, App.Path & "\netfiles\index.htm")
End Sub

Private Sub cmdRIGHT_Click()
    On Local Error Resume Next
    rtb.SelAlignment = 1
    rtb.SetFocus
    RaiseEvent Changed
End Sub

Private Sub cmdSPELL_Click()
    If spellChecker1.Visible Then Exit Sub

    If Len(rtb.Text) = 0 Then Exit Sub
    
    Call ED_Control(False)
    
    spellChecker1.Move (UserControl.Width - spellChecker1.Width) \ 2, (UserControl.Height - spellChecker1.Height) \ 2
    spellChecker1.Text = rtb.Text
    spellChecker1.Visible = True
    DoEvents
    spellChecker1.SpellCheck
End Sub

Private Sub spellChecker1_ChangeWord()
    Dim cWORD As String
    Dim pWORDPos As Long
    Dim nWORD As String
    
    Dim i As Long
    
    cWORD = Trim$(spellChecker1.Word)
    If Len(cWORD) = 0 Then Exit Sub
    
    pWORDPos = spellChecker1.WordPos
    nWORD = spellChecker1.NewWord
    
    
    ' the old word may be a capital letter
    ' at the start, if so capitalise the
    ' new word
    If Asc(Left$(cWORD, 1)) >= Asc("A") And Asc(Left$(cWORD, 1)) <= Asc("Z") Then
        nWORD = UCase$(Left$(nWORD, 1)) & Right$(nWORD, Len(nWORD) - 1)
    End If
    
    Focus_On_Word pWORDPos, cWORD, nWORD

    ' after word has changed update spell checker text
    spellChecker1.Text = rtb.Text
    
    spellChecker1.Get_Next_Word
End Sub

Private Sub spellChecker1_CurrentWord()
    Dim cWORD As String
    Dim pWORD As Long
    
    Dim i As Long
    
    cWORD = Trim$(spellChecker1.Word)
    If Len(cWORD) = 0 Then Exit Sub
    
    pWORD = spellChecker1.WordPos
    Focus_On_Word pWORD, cWORD
    
End Sub

Private Sub spellChecker1_Finished()
    On Local Error Resume Next
    spellChecker1.Visible = False
    Call ED_Control(True)
    rtb.SetFocus
End Sub
Private Sub cmdULINEOFF_Click()
    On Local Error Resume Next
    rtb.SelUnderline = False
    rtb.SetFocus
    RaiseEvent Changed
End Sub

Private Sub cmdULINEON_Click()
    On Local Error Resume Next
    rtb.SelUnderline = True
    rtb.SetFocus
    RaiseEvent Changed
End Sub

Private Sub Dir1_Change()
    Dim i As Integer
    On Local Error Resume Next
    File1.Path = Dir1.Path
    File1.Pattern = imgSPEC
    
    DoEvents
    List1.Clear
    For i = 0 To File1.ListCount - 1
        List1.AddItem File1.List(i)
    Next i
    Call AddScrollToListBox(UserControl.List1)
End Sub

Private Sub Drive1_Change()
    On Local Error Resume Next
    Dir1.Path = Drive1.Drive
    Call Dir1_Change
End Sub

Private Sub AddScrollToListBox(List As ListBox)
    Dim i As Integer
    Dim intGreatestLen As Integer
    Dim lngGreatestWidth As Long
    'Find Longest Text in Listbox


    For i = 0 To List.ListCount - 1


        If Len(List.List(i)) > Len(List.List(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i
    'Get Twips
    lngGreatestWidth = UserControl.TextWidth(List.List(intGreatestLen) + Space(4))
    'Space(1) is used to prevent the last Ch
    '     aracter from being cut off
    'Convert to Pixels
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    'Use api to add scrollbar
    SendMessage List.hWnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
    
End Sub

Private Sub File1_Click()
    Dim percRed As Double
    On Local Error Resume Next
    If Len(imgSPEC) = 0 Then Exit Sub
    Picture1.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)
    Image2.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)
    
    If Picture1.ScaleWidth <= MaxImageWidth Then
        htmlImageWidth = Picture1.ScaleWidth
        htmlImageHeight = Picture1.ScaleHeight
    Else
        ' need some re-sizing
        percRed = ((Picture1.ScaleWidth - MaxImageWidth) / Picture1.ScaleWidth) * 100
        
        htmlImageWidth = uMaxImageWidth
        htmlImageHeight = Picture1.ScaleHeight - ((percRed / 100) * Picture1.ScaleHeight)
    End If
    
    Image2.Width = Screen.TwipsPerPixelX * htmlImageWidth
    Image2.Height = Screen.TwipsPerPixelY * htmlImageHeight
        
   
    Image2.Left = 0
    Image2.Top = 0
    
    Picture2.Width = 3375
    Picture2.Height = 2775
    
    
    
    ' do scroll bars
    If Picture2.Width >= Image2.Width Then
        hs.Visible = False
    Else
        hs.Visible = True
        
        hs.Min = 0
        hs.Max = Image2.Width - Picture2.Width
        If hs.Max / 10 > 1 Then hs.LargeChange = hs.Max / 10
        If hs.Max / 50 > 1 Then hs.SmallChange = hs.Max / 50
        hs.Value = 0
    End If
    
    If Picture2.Height >= Image2.Height Then
        vs.Visible = False
    Else
        vs.Visible = True
        
        vs.Min = 0
        vs.Max = Image2.Height - Picture2.Height
        If vs.Max / 10 > 1 Then vs.LargeChange = vs.Max / 10
        If vs.Max / 50 > 1 Then vs.SmallChange = vs.Max / 50
        vs.Value = 0
    End If
        
    If hs.Visible = False And vs.Visible = False Then
        Picture2.Width = 3735
        Picture2.Height = 3135
    ElseIf hs.Visible = False Then
        Picture2.Height = 3135
    ElseIf vs.Visible = False Then
        Picture2.Width = 3735
    End If
    
    vs.Height = Picture2.Height
    hs.Width = Picture2.Width

End Sub

Private Sub File1_DblClick()
    Dim insIMG As String
    
    On Local Error Resume Next
    fNAME = Dir1.Path & "\" & File1.FileName
    
    txtImageLink.Text = Trim$(txtImageLink.Text)
    If Len(txtImageLink.Text) <> 0 Then
        insIMG = "<A HREF=" & Chr$(34) & txtImageLink.Text & Chr$(34) & ">"
    End If
    
    fNAME = Replace_Text(fNAME, "\", "/")
    fNAME = Replace_Text(fNAME, "//", "\\")
    
    insIMG = insIMG & "<IMG SRC=" & Chr$(34) & fNAME & Chr$(34) & " BORDER=0 WIDTH=" & htmlImageWidth & " HEIGHT=" & htmlImageHeight & ">"
    
    If Len(txtImageLink.Text) <> 0 Then
        insIMG = insIMG & "</A>"
    End If
    
    frameIMAGE.Visible = False
    
    Call ED_Control(True)
    
    rtb.SelStart = rtbSelstart
    rtb.SelText = insIMG
    Call rtb_Change
    rtb.SetFocus
End Sub

Private Sub Image1_Click()
    On Local Error Resume Next
    frameIMAGE.Visible = False
    
    Call ED_Control(True)
    
    rtb.SetFocus
End Sub


Private Sub List1_Click()
    On Local Error Resume Next
    File1.ListIndex = List1.ListIndex
    Call File1_Click
End Sub

Private Sub List1_DblClick()
    On Local Error Resume Next
    File1.ListIndex = List1.ListIndex
    Call File1_DblClick
End Sub

Private Sub rtb_Change()
    Dim hFN As Integer

    On Local Error Resume Next
        
    If uAutoGenHTML Then Call RTF_To_HTML
    
    If uShowPreview And PreviewValue = 1 Then
        If Len(htmPageName) = 0 Then
            htmPageName = App.Path & "\" & Format$(Date, "ddmmyyyy") & Format$(Time$, "HHMMSS") & UserControl.Ambient.DisplayName & ".htm"
        End If
        hFN = FreeFile
        Open htmPageName For Output As #hFN
        Print #hFN, "<html>"
        Print #hFN, "<head>"
        Print #hFN, "</head>"
        Print #hFN, "<body>"
        Print #hFN, rtfHTML
        Print #hFN, "</body>"
        Print #hFN, "</html>"
        Close #hFN

        web.Navigate htmPageName
    End If
    
    RaiseEvent Changed
End Sub

Private Sub rtb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then
        rtb.SelText = vbTab
        rtb.SetFocus
        RaiseEvent Changed
        KeyCode = 0
        Exit Sub
    End If
    
    If KeyCode = Asc("\") Then KeyCode = 0
    

End Sub

Private Sub rtb_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("\") Then KeyAscii = 0
    
End Sub

Private Sub UserControl_Initialize()
    Dim i As Integer
    
    With rtb
        .SelTabCount = 1
        .SelTabs(0) = 800 'change as you need
    End With
 
    On Local Error Resume Next
    For i = 10 To 36 Step 2
        cboFONTSIZE.AddItem i
    Next i
    
    cboFONTSIZE.ListIndex = 0
    
    uAutoGenHTML = True
    uMaxImageWidth = 500
    uPreviewMode = previewHalf
    uShowBorder = True
    uBackStyle = 1
    uShowFont = True
    uShowColor = True
    uShowBold = True
    uShowUline = True
    uShowAlignments = False
    uShowPreview = True
    uShowItalic = True
    uShowImage = False
    uShowHyperlink = True
    uBorderColor = vbBlack
    uBackColor = &H8000000F
    uEnabled = True
    uShowSpell = True
    uLoadAllFonts = False
    
    Call ShowButtons
    Call Draw_Borders
End Sub

Private Sub LoadFonts()
    Dim i As Integer
    
    On Local Error Resume Next

    cboFONTS.Clear
    If uLoadAllFonts = False Then
        cboFONTS.AddItem "MS Sans Serif"
        cboFONTS.AddItem "Arial"
        cboFONTS.AddItem "Tahoma"
        cboFONTS.AddItem "Verdana"
        cboFONTS.AddItem "Haettenschweiler"
        cboFONTS.AddItem "Times New Roman"
        cboFONTS.AddItem "Courier New"
        cboFONTS.AddItem "Georgia"
    Else
        For i = 0 To Screen.FontCount - 1  ' Determine number of fonts.
            cboFONTS.AddItem Screen.Fonts(i)  ' Put each font into list box.
        Next i
    End If
    
    cboFONTS.ListIndex = 0
End Sub

Public Function Add_Path_To_Image_Links(ByVal ImagePath As String, ByVal HTMLText As String) As String
    Dim i As Integer
    Dim img As String
    Dim a As Integer
    Dim B As Integer
    
    ImagePath = Replace_Text(ImagePath, "\", "/")
    
    i = InStr(1, HTMLText, "<IMG SRC=", vbTextCompare)
    While i <> 0
        a = InStr(i, HTMLText, Chr$(34), vbTextCompare)
        B = InStr(a + 1, HTMLText, Chr$(34), vbTextCompare)
        img = Mid$(HTMLText, a + 1, (B - a) - 1)
        If InStr(1, img, "\", vbTextCompare) = 0 Then
            HTMLText = Replace_Text(HTMLText, "<IMG SRC=" & Chr$(34) & img, "<IMG SRC=" & Chr$(34) & ImagePath & img)
        End If
        i = InStr(i + 1, HTMLText, "<IMG SRC=", vbTextCompare)
    Wend
    HTMLText = Replace_Text(HTMLText, "//", "\")
    Add_Path_To_Image_Links = HTMLText

End Function

Private Function GetFileName(pfName As String) As String
    Dim i As Integer
    
    On Local Error Resume Next

    For i = Len(pfName) To 1 Step -1
        If Mid$(pfName, i, 1) = "\" Then Exit For
        GetFileName = Mid$(pfName, i, 1) & GetFileName
    Next i
    
End Function

Private Function JustPath(Path As String) As String
    Dim Cnt As Integer
    
    On Local Error Resume Next
    
    Cnt = 1
    Do Until Mid(Path, Len(Path) - Cnt, 1) = "\"
        Cnt = Cnt + 1
    Loop
    JustPath = Left(Path, Len(Path) - Cnt)
End Function

Public Function Remove_Path_From_Image_Links(ByVal ImagePath As String, ByVal HTMLText As String) As String
    Dim i As Integer
    Dim img As String
    Dim a As Integer
    Dim B As Integer
    
    Dim newIMG As String
    Dim sourcePATH As String
    Dim sourceFNAME As String
    
    'ImagePath = Replace_Text(ImagePath, "\", "/")
    i = InStr(1, HTMLText, "<IMG SRC=", vbTextCompare)
    While i <> 0
        a = InStr(i, HTMLText, Chr$(34), vbTextCompare)
        B = InStr(a + 1, HTMLText, Chr$(34), vbTextCompare)
        img = Mid$(HTMLText, a + 1, (B - a) - 1)
        
        
        newIMG = Replace_Text(img, "/", "\")
        newIMG = LCase$(newIMG)
        sourcePATH = JustPath(newIMG)
        sourceFNAME = GetFileName(newIMG)
        
        ' if sourcepath is diff to imagepath(destination) then do not copy
        If sourcePATH <> ImagePath Then
            FileCopy newIMG, ImagePath & sourceFNAME
        End If
            
        
        If InStr(1, img, "\", vbTextCompare) = 0 Then
            HTMLText = Replace_Text(HTMLText, "<IMG SRC=" & Chr$(34) & img, "<IMG SRC=" & Chr$(34) & sourceFNAME)
        End If
        i = InStr(i + 1, HTMLText, "<IMG SRC=", vbTextCompare)
    Wend
    HTMLText = Replace_Text(HTMLText, "//", "\")
    Remove_Path_From_Image_Links = HTMLText

End Function

Private Sub Draw_Borders()
   
    On Local Error Resume Next
    
    UserControl.BackColor = uBackColor

    If uBackStyle = 0 Or uShowBorder = False Then
        Line1(0).Visible = False
        Line1(1).Visible = False
        Line1(2).Visible = False
        Line1(3).Visible = False
    Else
        Line1(0).Visible = True
        Line1(1).Visible = True
        Line1(2).Visible = True
        Line1(3).Visible = True
        
        ' top line
        LineSize 0, 0, 0, UserControl.ScaleWidth - 1, 0, uBorderColor

        ' left line
        LineSize 1, 0, 0, 0, UserControl.ScaleHeight - 1, uBorderColor

        ' right line
        LineSize 2, UserControl.ScaleWidth - 10, 0, UserControl.ScaleWidth - 10, UserControl.ScaleHeight, uBorderColor

        ' bottom line
        LineSize 3, 0, UserControl.ScaleHeight - 10, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10, uBorderColor
    End If
    
End Sub

Public Sub LineSize(i As Integer, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lcolor As Long)
    On Local Error Resume Next
    
    Line1(i).BorderColor = lcolor
    Line1(i).X1 = X1
    Line1(i).Y1 = Y1
    Line1(i).X2 = X2
    Line1(i).Y2 = Y2
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    
    If DoNotResize Then Exit Sub
    
    If UserControl.Width < 4635 Then UserControl.Width = 4635
    If UserControl.Height < 2940 Then UserControl.Height = 2940
      
    Call ShowButtons
    Call Draw_Borders
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        uAutoGenHTML = .ReadProperty("AutoGenerateHTML", True)
        uShowColor = .ReadProperty("ShowColor", True)
        uShowBold = .ReadProperty("ShowBold", True)
        uShowUline = .ReadProperty("ShowUline", True)
        uShowFont = .ReadProperty("ShowFont", True)
        uShowAlignments = .ReadProperty("ShowAlignments", True)
        uShowHyperlink = .ReadProperty("ShowHyperlink", True)
        uShowPreview = .ReadProperty("ShowPreview", True)
        uShowBorder = .ReadProperty("ShowBorder", True)
        uMaxImageWidth = .ReadProperty("MaxImageWidth", 500)
        uShowImage = .ReadProperty("ShowInsertImage", True)
        uBorderColor = .ReadProperty("BorderColor", vbBlack)
        uBackColor = .ReadProperty("BackColor", &H8000000F)
        uBackStyle = .ReadProperty("BackStyle", 1)
        uShowItalic = .ReadProperty("ShowItalic", True)
        uPreviewMode = .ReadProperty("PreviewMode", previewHalf)
        uEnabled = .ReadProperty("Enabled", True)
        uShowSpell = .ReadProperty("ShowSpellChecker", True)
        uLoadAllFonts = .ReadProperty("LoadAllFonts", False)
    End With
    UserControl.BackStyle = uBackStyle
    Call LoadFonts
    Call ShowButtons
    Call Draw_Borders
    Call ED_Control(uEnabled)
End Sub

Private Sub UserControl_Terminate()
    On Local Error Resume Next
    
    If Len(htmPageName) <> 0 Then
        If Len(Dir$(htmPageName)) <> 0 Then Kill htmPageName
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ShowColor", uShowColor, True
        .WriteProperty "ShowBold", uShowBold, True
        .WriteProperty "ShowUline", uShowUline, True
        .WriteProperty "ShowFont", uShowFont, True
        .WriteProperty "ShowAlignments", uShowAlignments, True
        .WriteProperty "ShowHyperlink", uShowHyperlink, True
        .WriteProperty "ShowPreview", uShowPreview, True
        .WriteProperty "ShowBorder", uShowBorder, True
        .WriteProperty "MaxImageWidth", uMaxImageWidth, 500
        .WriteProperty "ShowInsertImage", uShowImage, True
        .WriteProperty "BorderColor", uBorderColor, vbBlack
        .WriteProperty "BackColor", uBackColor, &H8000000F
        .WriteProperty "BackStyle", uBackStyle, 1
        .WriteProperty "ShowItalic", uShowItalic, True
        .WriteProperty "PreviewMode", uPreviewMode, previewHalf
        .WriteProperty "Enabled", uEnabled, True
        .WriteProperty "AutoGenerateHTML", uAutoGenHTML, True
        .WriteProperty "ShowSpellChecker", uShowSpell, True
        .WriteProperty "LoadAllFonts", uLoadAllFonts, False
    End With
End Sub

Private Sub hs_Change()
    On Local Error Resume Next
    Image2.Left = 0 - hs.Value
End Sub

Private Sub hs_Scroll()
    On Local Error Resume Next
    Image2.Left = 0 - hs.Value
End Sub

Private Sub vs_Change()
    On Local Error Resume Next
    Image2.Top = 0 - vs.Value
End Sub

Private Sub vs_Scroll()
    On Local Error Resume Next
    Image2.Top = 0 - vs.Value
End Sub

