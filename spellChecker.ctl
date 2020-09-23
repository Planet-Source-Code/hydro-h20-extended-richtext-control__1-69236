VERSION 5.00
Begin VB.UserControl spellChecker 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ClipControls    =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5760
   Begin VB.Frame Frame4 
      Height          =   4095
      Left            =   1200
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   4215
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   3120
         Width           =   3975
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   180
         Left            =   240
         Top             =   2880
         Width           =   3735
      End
      Begin VB.Shape Shape4 
         FillStyle       =   0  'Solid
         Height          =   180
         Left            =   240
         Top             =   2880
         Width           =   3735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Please Wait"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Initialising Word Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Change To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   4215
      Begin VB.TextBox txtCHANGETO 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdCHANGETO 
         Caption         =   "Change To"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   200
         Width           =   975
      End
   End
   Begin VB.ListBox soundExWords 
      Height          =   1815
      Left            =   6600
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Words Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4215
      Begin VB.CommandButton cmdIGNORE 
         Caption         =   "Ignore"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   13
         ToolTipText     =   "Ignore Word"
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdCHANGE 
         Caption         =   "Change"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCHANGEALL 
         Caption         =   "Change All"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdIGNOREALL 
         Caption         =   "Ignore All"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         ToolTipText     =   "Ignore ALL Words"
         Top             =   1920
         Width           =   975
      End
      Begin VB.ListBox levenWords 
         Height          =   2010
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Checking Word"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.CommandButton cmdADD 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Add Word To Dictionary"
         Top             =   200
         Width           =   975
      End
      Begin VB.Image imgALT 
         Height          =   600
         Left            =   3960
         Picture         =   "spellChecker.ctx":0000
         Stretch         =   -1  'True
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   60
         Left            =   120
         Top             =   555
         Width           =   2775
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   60
         Left            =   120
         Top             =   550
         Width           =   2775
      End
      Begin VB.Label lblWORD 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   4080
      Picture         =   "spellChecker.ctx":0442
      ToolTipText     =   "Close Spell Checker"
      Top             =   40
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "All Matches:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "spellChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Declare Function GetInputState Lib "user32.dll" () As Long

Private Type ChangeALLType
    OriginalWord As String
    ReplaceWord As String
End Type

Dim ChangeAllWords() As ChangeALLType
Dim IgnoreAllWords() As String

Public Event Error()
Public Event StatusChange()
Public Event CurrentWord()
Public Event ChangeWord()
Public Event Finished()
Public Event Alternate()

Dim WSP As Workspace
Dim wordDBS As Database
Dim wRST As Recordset

Dim uShowEachWord As Boolean
Dim wordCHANGED As Boolean
Dim uShowAlternate As Boolean
Dim uError As String
Dim uStatus As String
Dim uText As String
Dim uTextOriginal As String
Dim uWord As String
Dim nWORD As String
Dim nPOS As Long
Dim uWordPos As Long
Dim CloseClicked As Boolean

Private Sub RefreshAlternate()
    If uShowAlternate Then
        cmdADD.Left = 3000
        cmdADD.Width = 855
        'imgALT.Left = 3960
        imgALT.Visible = True
    Else
        cmdADD.Left = 3120
        cmdADD.Width = 975
        imgALT.Visible = False
        'imgALT.Left = -300
    End If
End Sub

Property Let ShowEachWord(nShow As Boolean)
    uShowEachWord = nShow
End Property

Property Get ShowEachWord() As Boolean
    ShowEachWord = uShowEachWord
End Property
Property Let ShowAlternate(nAlt As Boolean)
    uShowAlternate = nAlt
    Call RefreshAlternate
End Property

Property Get ShowAlternate() As Boolean
    ShowAlternate = uShowAlternate
End Property

Property Get WordPos() As Long
    WordPos = uWordPos
End Property

Public Sub SpellCheck()
    Dim odStatus As String
    
    uText = Trim$(uText)
    If Len(uText) = 0 Then
        uError = "Nothing To Check"
        RaiseEvent Error
        RaiseEvent Finished
        Exit Sub
    End If
    
    If Len(Dir$(App.Path & "\words.mdb")) = 0 Then
        If Len(Dir$(App.Path & "\words.txt")) = 0 Then
            uError = "Dictionary Cannot Be Created"
            RaiseEvent Error
            RaiseEvent Finished
        End If
    End If
    
    nPOS = 1
    uWordPos = 0
    uWord = ""
    nWORD = ""
    wordCHANGED = False
    CloseClicked = False
    ReDim ChangeAllWords(0 To 0)
    ReDim IgnoreAllWords(0 To 0)
    
    odStatus = Open_WordDB
    If odStatus <> "Ok" Then
        uError = odStatus
        RaiseEvent Error
        RaiseEvent Finished
    End If
    
    Call Get_Next_Word
End Sub

Public Sub Get_Next_Word()
    Dim i As Long
    Dim ii As Long
    Dim a As Integer
    Dim pPERC As Integer
    Dim pWIDTH As Long
    
    On Local Error Resume Next
    
    cmdADD.Enabled = False
    cmdCHANGE.Enabled = False
    cmdIGNORE.Enabled = False
    cmdIGNOREALL.Enabled = False
    cmdCHANGEALL.Enabled = False
    cmdCHANGETO.Enabled = False
    
nextWord:
    DoEvents
    
    pPERC = (nPOS / Len(uText)) * 100
    If pPERC > 100 Then pPERC = 100
    pWIDTH = (pPERC / 100) * Shape1.Width
    If pWIDTH > Shape1.Width Then pWIDTH = Shape1.Width
    Shape2.Width = pWIDTH
    Shape2.Refresh
    
    If CloseClicked Then
        wordDBS.Close
        Set wordDBS = Nothing
        WSP.Close
        Set WSP = Nothing
        
        RaiseEvent Finished
        Exit Sub
    End If
    
    If wordCHANGED Then
        ' if a word has changed, then the text will
        ' have been updated. The current nPOS may be in the
        ' middle of the text now.
        ' so move nPOS back to a space then add to
moveBACK:
        If nPOS = 0 Then GoTo moveEND
        If Mid$(uText, nPOS, 1) = "¬" Then GoTo moveEND
        nPOS = nPOS - 1
        GoTo moveBACK
'        While nPOS > 0 And Mid$(uText, nPOS, 1) <> " "
'            nPOS = nPOS - 1
'        Wend
moveEND:
        nPOS = nPOS + 1
    End If
    
    wordCHANGED = False
    i = InStr(nPOS, uText & " ", " ", vbTextCompare)
    ii = InStr(nPOS, uText, "¬", vbTextCompare)
    If i = 0 And ii = 0 Then
        wRST.Close
        Set wRST = Nothing
        wordDBS.Close
        Set wordDBS = Nothing
        WSP.Close
        Set WSP = Nothing
    
        RaiseEvent Finished
        Exit Sub
    End If
    
    If ii < i And ii <> 0 Then i = ii
    If i = 0 Then
        wRST.Close
        Set wRST = Nothing
        wordDBS.Close
        Set wordDBS = Nothing
        WSP.Close
        Set WSP = Nothing
        
        RaiseEvent Finished
        Exit Sub
    End If
    
    uWord = Mid$(uText, nPOS, i - nPOS)
    
    uWord = Replace_Text(uWord, "¬", "")
    If Len(uWord) = 0 Then
        nPOS = i + 1
        GoTo nextWord
    End If
    
RemoveQuotesETC:
    If Left$(uWord, 1) = "'" Or Left$(uWord, 1) = Chr$(34) Then
        uWord = Right$(uWord, Len(uWord) - 1)
        GoTo RemoveQuotesETC
    End If
    If Right$(uWord, 1) = "'" Or Right$(uWord, 1) = Chr$(34) Then
        uWord = Left$(uWord, Len(uWord) - 1)
        GoTo RemoveQuotesETC
    End If
    If Right$(uWord, 1) = "?" Or Right$(uWord, 1) = "!" Or Right$(uWord, 1) = "," Or Right$(uWord, 1) = "." Or Right$(uWord, 1) = ";" Or Right$(uWord, 1) = ":" Then
        uWord = Left$(uWord, Len(uWord) - 1)
        GoTo RemoveQuotesETC
    End If
    
    If Len(uWord) = 0 Then
        nPOS = i + 1
        GoTo nextWord
    End If
    
    lblWORD = uWord
    lblWORD.Refresh
    txtCHANGETO.Text = uWord
    txtCHANGETO.Refresh
    cmdCHANGETO.ToolTipText = "Change Word To " & txtCHANGETO.Text
    
    uWordPos = nPOS
    If uShowEachWord Then RaiseEvent CurrentWord
    
    If Word_Exists(uWord) Then
        ' word exists by unique index
        nPOS = i + 1
        GoTo nextWord
    End If
    
    
    ' word not found
    ' maybe in ignore list
    If UBound(IgnoreAllWords) > 0 Then
        For a = 1 To UBound(IgnoreAllWords)
            If LCase$(IgnoreAllWords(a)) = LCase$(uWord) Then
                ' in ignore list
                nPOS = i + 1
                GoTo nextWord
            End If
        Next a
    End If
    

    
    Call Check_Word(uWord)
    DoEvents
    If CloseClicked Then
        wRST.Close
        Set wRST = Nothing
        wordDBS.Close
        Set wordDBS = Nothing
        WSP.Close
        Set WSP = Nothing
        
        RaiseEvent Finished
        Exit Sub
    End If
    
    nPOS = i + 1
    
    If levenWords.ListCount = 1 Then
        If LCase$(levenWords.List(0)) = LCase$(uWord) Then GoTo nextWord
    End If
    
    ' word not found
    ' maybe in change all list
    If UBound(ChangeAllWords) > 0 Then
        For a = 1 To UBound(ChangeAllWords)
            If LCase$(ChangeAllWords(a).OriginalWord) = LCase$(uWord) Then
                RaiseEvent CurrentWord
                DoEvents
                
                For ii = 0 To levenWords.ListCount - 1
                    If LCase$(levenWords.List(ii)) = LCase$(ChangeAllWords(a).ReplaceWord) Then
                        levenWords.ListIndex = ii
                        Exit For
                    End If
                Next ii
            
                If ii > (levenWords.ListCount - 1) Then
                    ' hmm not in list so add to it
                    levenWords.AddItem ChangeAllWords(i).ReplaceWord
                    levenWords.ListIndex = levenWords.ListCount - 1
                End If
                
                nWORD = levenWords.List(levenWords.ListIndex)
                wordCHANGED = True
                RaiseEvent ChangeWord
           
                Exit Sub
            End If
        Next a
    End If
    
    levenWords.ListIndex = 0
    Call levenWords_Click

    RaiseEvent CurrentWord
    cmdADD.Enabled = True
    cmdCHANGE.Enabled = True
    cmdIGNORE.Enabled = True
    cmdIGNOREALL.Enabled = True
    cmdCHANGEALL.Enabled = True
    cmdCHANGETO.Enabled = True
End Sub

Property Let Text(nText As String)
    uText = nText
    uTextOriginal = uText
    uText = Replace_Text(uText, vbCrLf, "¬¬")
    uText = Replace_Text(uText, Chr$(9), "¬")
    
End Property

Property Get Word() As String
    Word = uWord
End Property

Property Get NewWord() As String
    NewWord = nWORD
End Property

Property Get Status() As String
    Status = uStatus
End Property

Property Get ErrorMessage() As String
    ErrorMessage = uError
End Property

Public Sub InitialiseWORDS()
    Dim r As String
    
    uStatus = "Initialising Words"
    RaiseEvent StatusChange
    
    r = Create_Word_Database
    
    If r <> "OK" Then
        uError = r
        uStatus = "Failed To Initialise Words"
        RaiseEvent Error
        RaiseEvent StatusChange
    Else
        uError = ""
        RaiseEvent StatusChange
    End If
    
End Sub

Private Function Open_WordDB() As String
    Dim wDB As String
    
    On Local Error Resume Next
    
    If Right$(App.Path, 1) = "\" Then
        wDB = App.Path & "Words.mdb"
    Else
        wDB = App.Path & "\Words.mdb"
    End If
    
    If Len(Dir$(wDB)) = 0 Then
        If Len(Dir$(App.Path & "\words.txt")) = 0 Then
            Open_WordDB = "Cannot Find Words Database"
            Exit Function
        Else
            Call InitialiseWORDS
        End If
    End If
    
    Set WSP = DBEngine.Workspaces(0)
    
    ' database created with not errors
    ' open database and record set
    Set wordDBS = WSP.OpenDatabase(wDB, False)

    Set wRST = wordDBS.OpenRecordset("Words", dbOpenTable)
    wRST.Index = "Word"

    Open_WordDB = "Ok"

End Function

Private Function Word_Exists(iWord As String) As Boolean
    On Local Error Resume Next
    
    wRST.Seek "=", iWord
    If wRST.NoMatch = False Then Word_Exists = True
End Function

Private Sub Check_Word(ByVal strWord As String)
    On Local Error GoTo subFAIL
    
    Dim SndxMatchRS As Recordset
    Dim LdMax As Long
    Dim lenTmp As Long
    Dim cPhoneme As New clsPhoneme
    Dim Soundex As String
    Dim LD As Long
    Dim i As Long
    Dim threshold As Long
    Dim strMATCH As String
    
    
    ' ensure word to search on
    strWord = Trim$(strWord)
    If strWord = vbNullString Then
        uError = "No Word To Search On"
        RaiseEvent Error
        Set cPhoneme = Nothing
        Exit Sub
    End If
    
    
    uStatus = "Searching..."
    RaiseEvent StatusChange
    
    
    '// Get the soundex of the input word
    Soundex = cPhoneme.GetSoundexWord(strWord)
        
        
    '// Now find all entries in the database which match the soundex of the input word
    Set SndxMatchRS = wordDBS.OpenRecordset("SELECT [word] from Words WHERE " & _
                                               "Soundex = " & _
                                               Chr$(34) & Soundex & Chr$(34), _
                                               dbOpenSnapshot)
                                
    '// Populate the Listbox (soundEXWords)
    soundExWords.Clear
    levenWords.Clear
    
    With SndxMatchRS
        While .EOF = False
            If GetInputState <> 0 Then DoEvents
            soundExWords.AddItem !Word
            lenTmp = Len(!Word)
            If lenTmp > LdMax Then LdMax = lenTmp
            .MoveNext
        Wend
    End With
    
    ' if no words in soundex list then will not find any words in leven list
    If soundExWords.ListCount = 0 Then
        ' no word matches
        uStatus = "No Words Match"
        GoTo subEXIT
    End If
    
    
    ' have filled main soundex list
    ' fill leven list
    threshold = 0
    strWord = UCase$(strWord)
    
ReDO:
    '// walk through all soundex matches
    For i = 0 To soundExWords.ListCount
        strMATCH = Trim$(soundExWords.List(i))
        If strMATCH <> vbNullString Then
            If GetInputState <> 0 Then DoEvents
            LD = cPhoneme.GetLevenshteinDistance(strWord, UCase$(strMATCH))
            
            '// Get all Levenshtein distances less than the scroll(threshold) value
            If LD <= threshold Then
                If LD < levenWords.ListCount Then
                    '// Add better matches up
                    levenWords.AddItem strMATCH, LD
                Else
                    levenWords.AddItem strMATCH
                End If
            End If
        End If
    Next i
    
    If levenWords.ListCount = 0 Then
        If threshold < LdMax Then
            threshold = threshold + 1
            GoTo ReDO
        End If
    End If
    
   
    uStatus = "Search Complete"
    GoTo subEXIT
    
subFAIL:
    uError = Err.Description & " (" & Err.Number & ")"
    uStatus = "Search FAILED"
    RaiseEvent Error
    
subEXIT:
    On Local Error Resume Next
    Set cPhoneme = Nothing
    
    SndxMatchRS.Close
    Set SndxMatchRS = Nothing
        
    RaiseEvent StatusChange
End Sub

Private Function Create_Word_Database() As String
    On Local Error GoTo funcFAIL
    
    Dim frf As Integer
    Dim tmpStr As String
    Dim i As Long
    Dim wordlistDB As Database
    Dim wordRS As Recordset
    Dim wDB As String
    Dim wWL As String
    Dim wlSIZE As Long
    Dim wlPOS As Long
    Dim wlPERC As Integer
    Dim cPhoneme As New clsPhoneme
    Dim RetDB As Boolean
    Dim sW As Long
    
    If Right$(App.Path, 1) = "\" Then
        wDB = App.Path & "Words.mdb"
        wWL = App.Path & "words.txt"
    Else
        wDB = App.Path & "\Words.mdb"
        wWL = App.Path & "\words.txt"
    End If
        
    ' if word database exits then just exit
    If Len(Dir$(wDB)) <> 0 Then
        ' word list already created
        Create_Word_Database = "OK"
        uStatus = "Already Initialised"
        Set cPhoneme = Nothing
        Frame4.Visible = False
        Exit Function
    End If

    Shape3.Width = 0
    Frame4.Move 120, 0
    Frame4.Visible = True
    Frame4.Refresh
    
    ' ensure the text file which is use to create database exists
    If Len(Dir$(wWL)) = 0 Then
        Create_Word_Database = "Cannot Find Word List To Import"
        Set cPhoneme = Nothing
        Frame4.Visible = False
        Exit Function
    End If
    

    ' create the database
    uStatus = "Creating Word Database"
    RaiseEvent StatusChange
    If GetInputState <> 0 Then DoEvents
    
    Set WSP = DBEngine.Workspaces(0)
    
    ' create a blank database
    RetDB = Create_New_Database(wDB, "", True)
    If RetDB = False Then
        Create_Word_Database = "Failed To Create Word Database"
        WSP.Close
        Set WSP = Nothing
        Set cPhoneme = Nothing
        Frame4.Visible = False
        Exit Function
    End If
    If GetInputState <> 0 Then DoEvents
    
    ' ceate blank table
    RetDB = Create_New_Table(wDB, "", "Words", True)
    If RetDB = False Then
        Create_Word_Database = "Failed To Create Words Table In Word Database"
        WSP.Close
        Set WSP = Nothing
        Set cPhoneme = Nothing
        Frame4.Visible = False
        Exit Function
    End If
    If GetInputState <> 0 Then DoEvents
    
    ' creat word field
    RetDB = Create_New_Field(wDB, "", "Words", "Word", dbText, True, True, False, True, Null, Null, Null, 50, True)
    If RetDB = False Then
        Create_Word_Database = "Failed To Create Word Field In Words Table"
        WSP.Close
        Set WSP = Nothing
        Set cPhoneme = Nothing
        Frame4.Visible = False
        Exit Function
    End If
    If GetInputState <> 0 Then DoEvents
    
    ' create soundex field
    RetDB = Create_New_Field(wDB, "", "Words", "Soundex", dbText, True, False, False, True, Null, Null, Null, 8, True)
    If RetDB = False Then
        Create_Word_Database = "Failed To Create Soundex Field In Words Table"
        WSP.Close
        Set WSP = Nothing
        Set cPhoneme = Nothing
        Frame4.Visible = False
        Exit Function
    End If
    If GetInputState <> 0 Then DoEvents
    
    ' create soundex field
    RetDB = Create_New_Field(wDB, "", "Words", "UserAdded", dbBoolean, False, False, False, True, Null, Null, Null, Null, True)
    If RetDB = False Then
        Create_Word_Database = "Failed To Create UserAdded Field In Words Table"
        WSP.Close
        Set WSP = Nothing
        Set cPhoneme = Nothing
        Frame4.Visible = False
        Exit Function
    End If
    If GetInputState <> 0 Then DoEvents
    
    ' database created with not errors
    ' open database and record set
    Set wordlistDB = WSP.OpenDatabase(wDB, False)
    
    ' open record set for word to be imported into
    Set wordRS = wordlistDB.OpenRecordset("Words", dbOpenTable)

    ' get file size so progress can be calculated
    wlSIZE = FileLen(wWL)

    ' get file handle for import
    frf = FreeFile()
    
    Open App.Path & "\words.txt" For Input As #frf
    
    
    With wordRS
        '// Read words from the file and add to the database
        Do While Not EOF(frf)
            If GetInputState <> 0 Then DoEvents
            Line Input #frf, tmpStr
            wlPOS = wlPOS + Len(tmpStr) + 2
            
            tmpStr = Trim$(tmpStr)
            .AddNew
            !Word = tmpStr
            !Soundex = cPhoneme.GetSoundexWord(tmpStr)
            .Update
            
            
            '// Prevent the UI from freezing up
            i = i + 1
            If i Mod 1000 = 0 Then
                wlPERC = Format$((wlPOS / wlSIZE) * 100, "0")
                If wlPERC > 100 Then wlPERC = 100
                uStatus = wlPERC & "% Imported"
                Label4 = uStatus
                Label4.Refresh
                RaiseEvent StatusChange
                
                ' do shape3 progress
                sW = (wlPERC / 100) * Shape4.Width
                If sW > Shape4.Width Then sW = Shape4.Width
                Shape3.Width = sW
                Shape3.Refresh
                
                DoEvents
            End If
        Loop
        .Close
    End With
    Close #frf

    uStatus = i & " Words Imported"
    
    Create_Word_Database = "OK"
    GoTo funcEXIT

funcFAIL:
    If Err.Number = 3022 Then
        wordRS.CancelUpdate
        Resume Next
    End If
    Create_Word_Database = Err.Description & " (" & Err.Number & ")"
    
funcEXIT:
    On Local Error Resume Next
    Set cPhoneme = Nothing
    
    wordRS.Close
    Set wordRS = Nothing
    
    wordlistDB.Close
    Set wordlistDB = Nothing
    
    WSP.Close
    Set WSP = Nothing
    
    Close #frf
    
    Frame4.Visible = False
End Function

Private Sub Add_Word(ByVal nWORD As String)
    On Local Error Resume Next
    
    Dim cPhoneme As New clsPhoneme
        
    With wRST
        .AddNew
        !Word = nWORD
        !Soundex = cPhoneme.GetSoundexWord(nWORD)
        !UserAdded = True
        .Update
    End With
    
    Set cPhoneme = Nothing
End Sub


Private Function Delete_Table(dBasePath As String, dBaseName As String, tblToDelete As String) As Boolean
    On Local Error GoTo fcnFailed
              
    Dim DBToAppend As Database
    Dim a As Integer
    
    Set DBToAppend = WSP.OpenDatabase(dBasePath & dBaseName, False)
    DBToAppend.TableDefs.Refresh
    For a = 0 To DBToAppend.TableDefs.Count - 1
        If LCase$(DBToAppend.TableDefs(a).Name) = LCase$(tblToDelete) Then
            DBToAppend.TableDefs.Delete (tblToDelete)
            DBToAppend.TableDefs.Refresh
            Delete_Table = True
            DBToAppend.Close
            Set DBToAppend = Nothing
            Exit Function
        End If
    Next a
    
fcnFailed:
    Delete_Table = False
    DBToAppend.Close
    Set DBToAppend = Nothing
    Exit Function
End Function

Private Function Create_New_Table(dBasePath As String, dBaseName As String, tblName As String, ReCreateOnColl As Boolean) As Boolean
    On Local Error GoTo fcnFailed
              
    Dim DBToAppend As Database
    Dim NewTable As TableDef
    Dim NewField As Field
    Dim Ret1 As Boolean
    Dim a As Integer
    
    Set DBToAppend = WSP.OpenDatabase(dBasePath & dBaseName, False)
              
    For a = 0 To DBToAppend.TableDefs.Count - 1
        If LCase$(DBToAppend.TableDefs(a).Name) = LCase$(tblName) Then
            If ReCreateOnColl = True Then
                Ret1 = Delete_Table(dBasePath, dBaseName, tblName)
                If Ret1 = True Then
                    Exit For
                Else
                    Create_New_Table = False
                    DBToAppend.Close
                    Set DBToAppend = Nothing
                    Exit Function
                End If
            Else
                Create_New_Table = True
                DBToAppend.Close
                Set DBToAppend = Nothing
                Exit Function
            End If
        End If
    Next a
    
    Set NewTable = DBToAppend.CreateTableDef(tblName)
    Set NewField = NewTable.CreateField("StartField", dbText)
    NewTable.Fields.Append NewField
    DBToAppend.TableDefs.Append NewTable
    Create_New_Table = True
    Set NewTable = Nothing
    DBToAppend.Close
    Set DBToAppend = Nothing
    Ret1 = Delete_Field(dBasePath, dBaseName, tblName, "Startfield")
    Exit Function
    
fcnFailed:
    On Local Error Resume Next
    Create_New_Table = False
    Set NewTable = Nothing
    DBToAppend.Close
    Set DBToAppend = Nothing
    Exit Function
End Function

Private Function Delete_Field(dBasePath As String, dBase As String, tblName As String, fldName As String) As Boolean
    On Local Error GoTo fcnFailed
              
    Dim DBToAppend As Database
    Dim TblToAppend As Recordset
    Dim CheckIndex As Index
    Dim a As Integer
    
    Set DBToAppend = WSP.OpenDatabase(dBasePath & dBase, False)
    ' test for indexed
    DBToAppend.TableDefs.Refresh
    For a = 0 To DBToAppend.TableDefs(tblName).Indexes.Count - 1
        If DBToAppend.TableDefs(tblName).Indexes(a).Fields = "+" & fldName Then
            DBToAppend.TableDefs(tblName).Indexes.Delete (fldName)
            Exit For
        End If
    Next a
    DBToAppend.TableDefs.Refresh
    For a = 0 To DBToAppend.TableDefs(tblName).Fields.Count - 1
        If LCase$(DBToAppend.TableDefs(tblName).Fields(a).Name) = LCase$(fldName) Then
            DBToAppend.TableDefs(tblName).Fields.Delete (fldName)
            Delete_Field = True
            Set TblToAppend = Nothing
            Set DBToAppend = Nothing
            Exit Function
        End If
    Next a
    
fcnFailed:
    Set TblToAppend = Nothing
    DBToAppend.Close
    Set DBToAppend = Nothing
    Delete_Field = False
    Exit Function
End Function

Private Function Create_New_Database(dBasePath As String, dBaseName As String, ReCreateOnColl As Boolean) As Boolean
    On Local Error GoTo fcnFailed
    Dim NewDatabase As Database
    Dim l As String
    
    l = Dir$(dBasePath & dBaseName)
    If Len(l) <> 0 Then
        If ReCreateOnColl = True Then
            If Len(Dir$(dBasePath & dBaseName & "Bak")) > 0 Then
                Kill dBasePath & dBaseName & "Bak"
            End If
            Name dBasePath & dBaseName As dBasePath & dBaseName & "Bak"
        Else
            Create_New_Database = True
            Exit Function
        End If
    End If
    Set NewDatabase = WSP.CreateDatabase(dBasePath & dBaseName, dbLangGeneral)
    NewDatabase.Close
    Set NewDatabase = Nothing
    Create_New_Database = True
    Exit Function
    
fcnFailed:
    Set NewDatabase = Nothing
    Create_New_Database = False
    Exit Function
End Function

Private Function Create_New_Field(dBasePath As String, dBaseName As String, tblName As String, nFieldName As String, nFieldType As Long, nFieldIndexed As Boolean, nFieldUnique As Boolean, nFieldPrimary As Boolean, AlZLength As Boolean, nFieldPosition As Variant, nFieldAbutes As Variant, nFieldDefaultValue As Variant, nFieldSize As Variant, ReCreateOnColl As Boolean) As Boolean
    On Local Error GoTo fcnFailed
    
    Dim DBToAppend As Database
    Dim TblToAppend As TableDef
    Dim NewField As Field
    Dim Ret1 As Boolean
    Dim a As Integer
    Dim SqlQ1 As String
    
    Set DBToAppend = WSP.OpenDatabase(dBasePath & dBaseName, False)
    Set TblToAppend = DBToAppend.TableDefs(tblName)
    For a = 0 To DBToAppend.TableDefs(tblName).Fields.Count - 1
        If LCase$(DBToAppend.TableDefs(tblName).Fields(a).Name) = LCase$(nFieldName) Then
            If ReCreateOnColl = True Then
                Ret1 = Delete_Field(dBasePath, dBaseName, tblName, nFieldName)
                If Ret1 = False Then
                    Create_New_Field = False
                    Set NewField = Nothing
                    DBToAppend.Close
                    Set DBToAppend = Nothing
                    Exit Function
                End If
                Exit For
            Else
                Create_New_Field = True
                Set NewField = Nothing
                DBToAppend.Close
                Set DBToAppend = Nothing
                Exit Function
            End If
        End If
    Next a
    
    Set NewField = TblToAppend.CreateField(nFieldName, nFieldType)
    If Not IsNull(nFieldAbutes) Then
        NewField.Attributes = nFieldAbutes
    End If
    If Not IsNull(nFieldSize) Then
        NewField.Size = nFieldSize
    End If
    If Not IsNull(nFieldDefaultValue) Then
        NewField.DefaultValue = nFieldDefaultValue
    End If
    If Not IsNull(nFieldPosition) Then
        NewField.OrdinalPosition = nFieldPosition
    End If
    If AlZLength = True Then
        If nFieldType = dbText Or nFieldType = dbMemo Then
            NewField.AllowZeroLength = AlZLength
        End If
    End If
    TblToAppend.Fields.Append NewField
    SqlQ1 = ""
    If nFieldIndexed = True Then
        If nFieldUnique = True Then
            SqlQ1 = "CREATE UNIQUE INDEX [" & nFieldName
            SqlQ1 = SqlQ1 & "] ON " & tblName & "([" & nFieldName & "]);"
        Else
            SqlQ1 = "CREATE INDEX [" & nFieldName
            SqlQ1 = SqlQ1 & "] ON " & tblName & "([" & nFieldName & "]);"
        End If
    End If
    If nFieldPrimary = True Then
        SqlQ1 = "CREATE UNIQUE INDEX [" & nFieldName
        SqlQ1 = SqlQ1 & "] ON " & tblName & "([" & nFieldName & "]) WITH PRIMARY;"
    End If
    If Len(SqlQ1) <> 0 Then DBToAppend.Execute SqlQ1

    Create_New_Field = True
    DBToAppend.Close
    Set DBToAppend = Nothing
    Set NewField = Nothing
    Exit Function
    
fcnFailed:
    DBToAppend.Close
    Set DBToAppend = Nothing
    Set NewField = Nothing
    Create_New_Field = False
    Exit Function
End Function


Private Sub cmdADD_Click()
    Call Add_Word(lblWORD.Caption)
    Call Get_Next_Word
End Sub

Private Sub cmdCHANGE_Click()
    If levenWords.ListIndex = -1 Then Exit Sub
    
    nWORD = levenWords.List(levenWords.ListIndex)
    wordCHANGED = True
    RaiseEvent ChangeWord
    
End Sub

Private Sub cmdCHANGEALL_Click()
    Dim uITEM As Long
    
    On Local Error Resume Next
    
    If levenWords.ListIndex = -1 Then Exit Sub
    
    uITEM = UBound(ChangeAllWords)
    uITEM = uITEM + 1
    ReDim Preserve ChangeAllWords(0 To uITEM)
    
    ChangeAllWords(uITEM).OriginalWord = lblWORD
    ChangeAllWords(uITEM).ReplaceWord = levenWords.List(levenWords.ListIndex)
    
    Call cmdCHANGE_Click
End Sub


Private Sub cmdCHANGETO_Click()
    txtCHANGETO.Text = Trim$(txtCHANGETO.Text)
    If Len(txtCHANGETO.Text) = 0 Then Exit Sub
    If LCase$(txtCHANGETO.Text) = LCase$(lblWORD) Then Exit Sub
    
    nWORD = txtCHANGETO
    wordCHANGED = True
    RaiseEvent ChangeWord
End Sub

Private Sub cmdIGNORE_Click()
    Call Get_Next_Word
End Sub

Private Sub cmdIGNOREALL_Click()
    Dim uITEM As Long
    
    On Local Error Resume Next
    
    uITEM = UBound(IgnoreAllWords)
    uITEM = uITEM + 1
    ReDim Preserve IgnoreAllWords(0 To uITEM)
    
    IgnoreAllWords(uITEM) = lblWORD
    Call Get_Next_Word
End Sub

Private Sub Image1_Click()
    CloseClicked = True
    If cmdADD.Enabled Then
        wRST.Close
        Set wRST = Nothing
        wordDBS.Close
        Set wordDBS = Nothing
        WSP.Close
        Set WSP = Nothing
    
        RaiseEvent Finished
    End If
End Sub

Private Sub imgALT_Click()
    On Local Error Resume Next
    RaiseEvent Alternate
End Sub

Private Sub levenWords_Click()
    If levenWords.ListIndex = -1 Then Exit Sub
    cmdCHANGE.ToolTipText = "Change Word To " & levenWords.List(levenWords.ListIndex)
    cmdCHANGEALL.ToolTipText = "Change ALL Words To " & levenWords.List(levenWords.ListIndex)
End Sub

Private Sub txtCHANGETO_Change()
    cmdCHANGETO.ToolTipText = "Change Word To " & txtCHANGETO.Text
End Sub

Private Sub UserControl_Initialize()
    uShowAlternate = False
    Call RefreshAlternate
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 4455 '4440
    UserControl.Height = 4230 '3555
End Sub

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

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        uShowAlternate = .ReadProperty("ShowAlternate", False)
        uShowEachWord = .ReadProperty("ShowEachWord", False)
    End With
    Call RefreshAlternate
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ShowAlternate", uShowAlternate, False
        .WriteProperty "ShowEachWord", uShowEachWord, False
    End With
End Sub
