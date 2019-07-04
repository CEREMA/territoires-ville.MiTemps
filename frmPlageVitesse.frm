VERSION 5.00
Begin VB.Form frmPlageVitesse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modification des plages de vitesses"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmPlageVitesse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameValeurs 
      Caption         =   "Valeurs Début et Fin :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   1200
      TabIndex        =   21
      Top             =   120
      Width           =   2175
      Begin VB.ComboBox ComboVal6 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2100
         Width           =   735
      End
      Begin VB.ComboBox ComboVal5 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1740
         Width           =   735
      End
      Begin VB.ComboBox ComboVal4 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1380
         Width           =   735
      End
      Begin VB.ComboBox ComboVal3 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1020
         Width           =   735
      End
      Begin VB.ComboBox ComboVal2 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   660
         Width           =   735
      End
      Begin VB.ComboBox ComboVal1 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   735
      End
      Begin VB.Label LabelVal6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "supérieure à "
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label LabelVal5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "entre 000 et "
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label LabelVal4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "entre 000 et "
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label LabelVal3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "entre 000 et "
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label LabelVal2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "entre 000 et "
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   900
      End
      Begin VB.Label LabelVal1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "entre 000 et "
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame FramePlage 
      Caption         =   "Plage :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
      Begin VB.Label LabelPlage6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plage 6"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label LabelPlage5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plage 5"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label LabelPlage4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plage 4"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label LabelPlage3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plage 3"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label LabelPlage2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plage 2"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   540
      End
      Begin VB.Label LabelPlage1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plage 1"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmPlageVitesse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
