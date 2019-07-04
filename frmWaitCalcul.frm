VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmWaitCalcul 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Progression du calcul"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmWaitCalcul.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label LabelN 
      AutoSize        =   -1  'True
      Caption         =   "N%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   285
   End
   Begin VB.Label Label100 
      AutoSize        =   -1  'True
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7200
      TabIndex        =   2
      Top             =   360
      Width           =   465
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8,25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "frmWaitCalcul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CentrerFenetreEcran Me
End Sub
