VERSION 5.00
Begin VB.Form frmSRTBTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Richtext Box With Spell Checker"
   ClientHeight    =   10365
   ClientLeft      =   2610
   ClientTop       =   1545
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   10350
   Begin VB.Frame Frame1 
      Caption         =   "HTML"
      Height          =   2415
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   7800
      Width           =   9855
      Begin VB.TextBox Text2 
         Height          =   2055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rich Text"
      Height          =   2415
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   5280
      Width           =   9855
      Begin VB.TextBox Text1 
         Height          =   2055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   9615
      End
   End
   Begin superRTBTest.superRTB superRTB1 
      Height          =   5100
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8996
      ShowBorder      =   0   'False
   End
End
Attribute VB_Name = "frmSRTBTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub superRTB1_Changed()
Text1 = superRTB1.TextRTF
Text2 = superRTB1.TextHTML
End Sub
