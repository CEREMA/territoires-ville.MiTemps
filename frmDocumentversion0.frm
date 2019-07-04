VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mschrt20.ocx"
Begin VB.Form frmDocument0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmDocument0"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15030
   Icon            =   "frmDocument.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   15030
   Begin VB.PictureBox PictureItiRef 
      AutoRedraw      =   -1  'True
      Height          =   7320
      Left            =   13920
      ScaleHeight     =   7260
      ScaleWidth      =   660
      TabIndex        =   12
      Top             =   0
      Width           =   720
      Begin VB.Label LabelM�tre 
         AutoSize        =   -1  'True
         Caption         =   "(m)"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   210
      End
      Begin VB.Line AxeDist 
         X1              =   90
         X2              =   90
         Y1              =   7080
         Y2              =   240
      End
      Begin VB.Line Line2 
         X1              =   90
         X2              =   150
         Y1              =   240
         Y2              =   420
      End
      Begin VB.Line Line3 
         X1              =   90
         X2              =   30
         Y1              =   240
         Y2              =   420
      End
      Begin VB.Label LabelRep 
         AutoSize        =   -1  'True
         Caption         =   "Distance"
         Height          =   195
         Left            =   10
         TabIndex        =   13
         Top             =   60
         Width           =   630
      End
      Begin VB.Image ImageRepere 
         Height          =   480
         Index           =   0
         Left            =   150
         Picture         =   "frmDocument.frx":0442
         Stretch         =   -1  'True
         ToolTipText     =   "Autre rep�re"
         Top             =   1440
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin TabDlg.SSTab TabData 
      Height          =   7320
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   12912
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Itin�raire de R�f�rence"
      TabPicture(0)   =   "frmDocument.frx":0884
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LabelNomIti"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LabelLongIti"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LabelM�tres"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LabelInfoColor"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TextNomIti"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TextLongIti"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "BtnAjoutRep"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "BtnSuppRep"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CheckSection"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ComboRepFinSec"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "BtnMeteo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "BtnFiltreSel"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ComboRepDebSec"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "BtnSuppPar"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "SpreadParcours"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "SpreadRepere"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Courbe Distance/Temps"
      TabPicture(1)   =   "frmDocument.frx":08A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SpreadInfoParcoursDT"
      Tab(1).Control(1)=   "PicBoxDT"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Courbe Distance/Vitesse"
      TabPicture(2)   =   "frmDocument.frx":08BC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SpreadInfoParcoursDV"
      Tab(2).Control(1)=   "PicBoxDV"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Synoptique des Vitesses"
      TabPicture(3)   =   "frmDocument.frx":08D8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameLegende"
      Tab(3).Control(1)=   "PicBoxSynoV"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Histogramme des Vitesses"
      TabPicture(4)   =   "frmDocument.frx":08F4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "MSChart1"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Tableau Brut"
      TabPicture(5)   =   "frmDocument.frx":0910
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "SpreadTabBrut"
      Tab(5).Control(1)=   "BtnExportTabBrut"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Synth�se / Statistiques"
      TabPicture(6)   =   "frmDocument.frx":092C
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "SpreadTabSS"
      Tab(6).Control(1)=   "BtnExportTabSS"
      Tab(6).ControlCount=   2
      Begin FPSpread.vaSpread SpreadTabSS 
         Height          =   4815
         Left            =   -73170
         TabIndex        =   36
         Top             =   1920
         Width           =   10700
         _Version        =   131077
         _ExtentX        =   18874
         _ExtentY        =   8493
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         OperationMode   =   1
         ProcessTab      =   -1  'True
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmDocument.frx":0948
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread SpreadTabBrut 
         Height          =   4815
         Left            =   -73080
         TabIndex        =   34
         Top             =   1800
         Width           =   10570
         _Version        =   131077
         _ExtentX        =   18644
         _ExtentY        =   8493
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         OperationMode   =   1
         ProcessTab      =   -1  'True
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frmDocument.frx":0EFC
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin FPSpread.vaSpread SpreadRepere 
         Height          =   2295
         Left            =   60
         TabIndex        =   6
         Top             =   1440
         Width           =   10605
         _Version        =   131077
         _ExtentX        =   18706
         _ExtentY        =   4048
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   4
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         MaxRows         =   10
         ProcessTab      =   -1  'True
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   10
         SpreadDesigner  =   "frmDocument.frx":1349
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   5
         VisibleRows     =   5
      End
      Begin FPSpread.vaSpread SpreadParcours 
         Height          =   2055
         Left            =   60
         TabIndex        =   7
         Top             =   4320
         Width           =   11205
         _Version        =   131077
         _ExtentX        =   19764
         _ExtentY        =   3625
         _StockProps     =   64
         BackColorStyle  =   1
         EditEnterAction =   4
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   14
         MaxRows         =   15
         ProcessTab      =   -1  'True
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   10
         SpreadDesigner  =   "frmDocument.frx":251B
         StartingRowNumber=   0
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   14
         VisibleRows     =   15
      End
      Begin FPSpread.vaSpread SpreadInfoParcoursDT 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   11
         Top             =   600
         Width           =   2220
         _Version        =   131077
         _ExtentX        =   3916
         _ExtentY        =   9763
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayColHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   1
         MaxRows         =   13
         RowHeaderDisplay=   2
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   0
         SelectBlockOptions=   10
         SpreadDesigner  =   "frmDocument.frx":35B2
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   1
         VisibleRows     =   10
      End
      Begin FPSpread.vaSpread SpreadInfoParcoursDV 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   22
         Top             =   600
         Width           =   2220
         _Version        =   131077
         _ExtentX        =   3916
         _ExtentY        =   9763
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayColHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   1
         MaxRows         =   13
         RowHeaderDisplay=   2
         ScrollBarExtMode=   -1  'True
         ScrollBars      =   0
         SelectBlockOptions=   10
         SpreadDesigner  =   "frmDocument.frx":3D4F
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   1
         VisibleRows     =   10
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   5415
         Left            =   -74400
         OleObjectBlob   =   "frmDocument.frx":44EC
         TabIndex        =   23
         Top             =   500
         Width           =   8000
      End
      Begin VB.CommandButton BtnSuppPar 
         Caption         =   "Supprimer un parcours"
         Height          =   345
         Left            =   3960
         TabIndex        =   38
         Top             =   3855
         Width           =   1815
      End
      Begin VB.ComboBox ComboRepDebSec 
         Height          =   315
         ItemData        =   "frmDocument.frx":6906
         Left            =   1360
         List            =   "frmDocument.frx":6908
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   990
         Width           =   2800
      End
      Begin VB.CommandButton BtnExportTabSS 
         Caption         =   "Exporter en fichier texte (*.txt) avec comme s�parateur le point virgule..."
         Height          =   495
         Left            =   -64320
         TabIndex        =   37
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton BtnExportTabBrut 
         Caption         =   "Exporter en fichier texte (*.txt) avec comme s�parateur le point virgule..."
         Height          =   495
         Left            =   -64320
         TabIndex        =   35
         Top             =   600
         Width           =   3015
      End
      Begin VB.Frame FrameLegende 
         Caption         =   "L�gende "
         Height          =   2895
         Left            =   -74880
         TabIndex        =   26
         Top             =   540
         Width           =   1335
         Begin VB.Label Label7 
            Caption         =   "V en km/h"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label LabelClassV6 
            AutoSize        =   -1  'True
            Caption         =   "V > 140"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   2520
            Width           =   555
         End
         Begin VB.Label LabelClassV5 
            AutoSize        =   -1  'True
            Caption         =   "130 < V <= 140"
            Height          =   195
            Left            =   90
            TabIndex        =   31
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label LabelClassV4 
            AutoSize        =   -1  'True
            Caption         =   "120 < V <= 130"
            Height          =   195
            Left            =   90
            TabIndex        =   30
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label LabelClassV3 
            AutoSize        =   -1  'True
            Caption         =   "110 < V <= 120"
            Height          =   195
            Left            =   90
            TabIndex        =   29
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label LabelClassV2 
            AutoSize        =   -1  'True
            Caption         =   "100 < V <= 110"
            Height          =   195
            Left            =   90
            TabIndex        =   28
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label LabelClassV1 
            AutoSize        =   -1  'True
            Caption         =   "0 <= V <= 100"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1005
         End
      End
      Begin VB.PictureBox PicBoxSynoV 
         AutoRedraw      =   -1  'True
         Height          =   6620
         Left            =   -73440
         ScaleHeight     =   6555
         ScaleWidth      =   12135
         TabIndex        =   24
         Top             =   600
         Width           =   12200
      End
      Begin VB.CommandButton BtnFiltreSel 
         Caption         =   "Filtre de s�lection..."
         Height          =   345
         Left            =   120
         TabIndex        =   21
         Top             =   3855
         Width           =   1815
      End
      Begin VB.CommandButton BtnMeteo 
         Caption         =   "Conditions M�t�o..."
         Height          =   345
         Left            =   2025
         TabIndex        =   20
         Top             =   3855
         Width           =   1815
      End
      Begin VB.ComboBox ComboRepFinSec 
         Height          =   315
         Left            =   4400
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   990
         Width           =   2800
      End
      Begin VB.CheckBox CheckSection 
         Caption         =   "Section de travail entre :"
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Top             =   920
         Width           =   1260
      End
      Begin VB.PictureBox PicBoxDV 
         AutoRedraw      =   -1  'True
         Height          =   6620
         Left            =   -72600
         ScaleHeight     =   6555
         ScaleWidth      =   11295
         TabIndex        =   15
         Top             =   600
         Width           =   11360
      End
      Begin VB.CommandButton BtnSuppRep 
         Caption         =   "Supprimer un rep�re"
         Height          =   345
         Left            =   9045
         TabIndex        =   10
         Top             =   975
         Width           =   1575
      End
      Begin VB.CommandButton BtnAjoutRep 
         Caption         =   "Cr�er un rep�re"
         Height          =   345
         Left            =   7340
         TabIndex        =   9
         Top             =   975
         Width           =   1575
      End
      Begin VB.TextBox TextLongIti 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   4
         Top             =   570
         Width           =   855
      End
      Begin VB.TextBox TextNomIti 
         Height          =   285
         Left            =   720
         MaxLength       =   100
         TabIndex        =   2
         Top             =   570
         Width           =   7815
      End
      Begin VB.PictureBox PicBoxDT 
         AutoRedraw      =   -1  'True
         Height          =   6660
         Left            =   -72600
         ScaleHeight     =   6600
         ScaleWidth      =   11295
         TabIndex        =   8
         Top             =   520
         Width           =   11360
      End
      Begin VB.Label LabelInfoColor 
         AutoSize        =   -1  'True
         Caption         =   "Double cliquer sur les couleurs pour les changer"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5895
         TabIndex        =   25
         Top             =   3960
         Width           =   3390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "et"
         Height          =   195
         Left            =   4200
         TabIndex        =   19
         Top             =   1060
         Width           =   135
      End
      Begin VB.Label LabelM�tres 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10470
         TabIndex        =   5
         Top             =   600
         Width           =   150
      End
      Begin VB.Label LabelLongIti 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Longueur : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8610
         TabIndex        =   3
         Top             =   600
         Width           =   990
      End
      Begin VB.Label LabelNomIti 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nom : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable stockant la liste des parcours affect�s
Public maColParcours As New ColParcours
'Variable s'incr�mentant � chaque ajout d'un parcours affect�
Public monNbParcours As Integer

'Tableau donnant si l'onglet num�ro i, i allant de 1 � 6 doit �tre
'recalcul ou redessiner lors de son activation car des modifs ont
'eu lieu dans l'onglet ItiRef
Private monTabRedOnglet(1 To 6) As Boolean

'Variable indiquant si une modif a eu lieu
Public maModif As Boolean

'Variable donnant la taille de la marge du haut en impression
Public maMargeHaut As Single

'Variable stockant les pas de graduation de niveau 1 et 2
Public monPasGrad1 As Long
Public monPasGrad2 As Long

'Collection stockant les libell�s des diff�rentes conditions m�t�o
Public maColMeteo As New Collection

'Variable stockant la liste des rep�res
Public maColRepere As New ColRepere
'Variable s'incr�mentant � chaque ajout d'un rep�re
Public monNbRepere As Integer

'Variables donnant l'indice du parcours s�lectionn�
'dans les courbes Distance/Temps et Distance/Vitesse
Public monIndParcoursSelectDT As Integer
Public monIndParcoursSelectDV As Integer

'Variables donnant le maximun des Distances �cran en X et Y
'Elles seront affect�es dans la fonction dessinercourbe si
'on ne fait pas d'impression
Public maDistMaxEcranX As Single
Public maDistMaxEcranY As Single

'Variables donnant le maximun �cran en Y et le minimum �cran en X
'Elles seront affect�es dans la fonction dessinercourbe si
Public monMinXecran As Single
Public monMaxYecran As Single

'Variables donnant les min et max des temps, distance et vitesse
'pour les courbes DT et DV
Public monMaxT As Single
Public monMaxV As Single
Public monMaxD As Single
Public monMinT As Single
Public monMinV As Single
Public monMinD As Single

'Variables stockant les min et max des distances en zoom tout
'sans section de travail
Public monMaxDtot As Single
Public monMinDtot As Single

'Variables stockant le nom de l'itin�raire et sa longueur
Public monNomIti As String
Public maLongIti As Single

'Variable stockant l'id de fichier mit
Public monFichId As Integer

'Variable stockant les modifications d'�chelle pour l'impression
'des courbes Distance/Temps, Distance/Vitesse, Synoptique des vitesses
Public maModifEchImp As Integer
Public monMinDImp As Single
Public monMaxDImp As Single
Public monMinTImp As Single
Public monMaxTImp As Single
Public monMinVImp As Single
Public monMaxVImp As Single


Private Sub BtnAjoutrep_Click()
    AfficherNouveauRepere monIti
End Sub

Private Sub BtnExportTabBrut_Click()
    Dim unNomFich As String, unFileId As Integer
    Dim unSep As String, uneLigneTexte As String, uneStrTmp As String
    
    unNomFich = frmMain.ChoisirFichier(MsgExportTxtAs, MsgTxtFile, CurDir)
    unNbTroncon = 0
    If unNomFich <> "" Then
        Me.MousePointer = vbHourglass
        unFileId = FreeFile(0)
        'Ouverture du fichier
        On Error GoTo ErreurOpenFileExpTB
        Open unNomFich For Output As #unFileId
        'Remplissage du fichier
        For i = 0 To SpreadTabBrut.MaxRows
            SpreadTabBrut.Row = i
            uneLigneTexte = ""
            For j = 1 To SpreadTabBrut.MaxCols
                SpreadTabBrut.Col = j
                If j = SpreadTabBrut.MaxCols Then
                    unSep = ""
                    uneStrTmp = ChangerCREnPtVirg(SpreadTabBrut.Text)
                ElseIf j = 1 And SpreadTabBrut.Text = "" Then
                    'On essaye d'avoir 4 colonnes vides
                    unSep = ";"
                    uneStrTmp = String(4, ";")
                ElseIf j = 1 And Mid(SpreadTabBrut.Text, 1, 2) = "De" Then
                    'On essaye d'avoir tron�on num�ro en 1�re colonne
                    unSep = ";"
                    unNbTroncon = unNbTroncon + 1
                    uneStrTmp = "Tron�on " + Format(unNbTroncon) + unSep + ChangerCREnPtVirg(SpreadTabBrut.Text)
                ElseIf i = 0 And j = 3 Then
                    'Modif du contenu de l'ent�te de la colonne 3
                    'Distance et temps de parcours avec 3 CR devient
                    'Distance de parcours;Temps de parcours
                    uneStrTmp = "Distance parcourue;Temps de parcours"
                Else
                    uneStrTmp = ChangerCREnPtVirg(SpreadTabBrut.Text)
                    unSep = ";"
                End If
                uneLigneTexte = uneLigneTexte + uneStrTmp + unSep
            Next j
            Print #unFileId, uneLigneTexte
        Next i
        'Fermeture du fichier
        Close #unFileId
        Me.MousePointer = vbDefault
        'Shell "notepad " + unNomFich, vbNormalFocus
        MsgBox "Fin de l'exportation du Tableau Brut dans " + unNomFich, vbInformation
        ' D�sactive la r�cup�ration d'erreur.
        On Error GoTo 0
    End If
    
    'Sortie pour �viter le passage dans la gestion d'erreur
    Exit Sub
    
ErreurOpenFileExpTB:
    ' D�sactive la r�cup�ration d'erreur.
    Me.MousePointer = vbDefault
    MsgBox "Erreur " + Format(Err.Number) + " : " + Err.Description + " (Fichier peut-�tre d�j� ouvert par une autre application)", vbCritical
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub BtnExportTabSS_Click()
    Dim unNomFich As String, unFileId As Integer
    Dim unSep As String, uneLigneTexte As String, uneStrTmp As String
    
    unNomFich = frmMain.ChoisirFichier(MsgExportTxtAs, MsgTxtFile, CurDir)
    If unNomFich <> "" Then
        Me.MousePointer = vbHourglass
        unFileId = FreeFile(0)
        'Ouverture du fichier
        On Error GoTo ErreurOpenFileExpTSS
        Open unNomFich For Output As #unFileId
        'Remplissage du fichier
        unSep = ";"
        For i = 0 To SpreadTabSS.MaxRows
            SpreadTabSS.Row = i
            uneLigneTexte = ""
            For j = 1 To SpreadTabSS.MaxCols
                SpreadTabSS.Col = j
                If j = 1 Then
                    If i = 0 Then
                        uneLigneTexte = "Information Tron�on" + unSep
                    ElseIf SpreadTabSS.Text = "" Then
                        uneLigneTexte = unSep
                    Else
                        uneLigneTexte = uneLigneTexte + SpreadTabSS.Text + unSep
                    End If
                ElseIf j = SpreadTabSS.MaxCols Then
                    uneLigneTexte = uneLigneTexte + SpreadTabSS.Text
                Else
                    uneLigneTexte = uneLigneTexte + SpreadTabSS.Text + unSep
                End If
            Next j
            Print #unFileId, uneLigneTexte
        Next i
        'Fermeture du fichier
        Close #unFileId
        Me.MousePointer = vbDefault
        MsgBox "Fin de l'exportation du Tableau Synth�se/Statistique dans " + unNomFich, vbInformation
        ' D�sactive la r�cup�ration d'erreur.
        On Error GoTo 0
    End If
    
    'Sortie pour �viter le passage dans la gestion d'erreur
    Exit Sub
    
ErreurOpenFileExpTSS:
    ' D�sactive la r�cup�ration d'erreur.
    Me.MousePointer = vbDefault
    MsgBox "Erreur " + Format(Err.Number) + " : " + Err.Description + " (Fichier peut-�tre d�j� ouvert par une autre application)", vbCritical
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub BtnFiltreSel_Click()
    Dim unY1 As Long, unY2 As Long
    
    frmFiltreSel.Show vbModal
    
    If monBtnClick = "OK" Then
        'Remplir la colonne deux du spread parcours de l'itin�raire actif
        'On masque le spreadparcours car le calcul du parcours est fait �
        'cochage du champ utilis� sauf s'il n'est pas visible (cf la fonction
        'spreadparcours_buttonclicked de ce module)
        SpreadParcours.Visible = False
        For i = 1 To SpreadParcours.MaxRows
            SpreadParcours.Row = i
            SpreadParcours.Col = 2
            SpreadParcours.Value = Abs(maColParcours(i).monIsUtil)
        Next i
        SpreadParcours.Visible = True
        'Affichage pour calcul du parcours moyen
        If CheckSection.Value = 0 Then
            'Stockage des abs d�but et fin du parcours
            unY1 = -100
            unY2 = 1000000
        Else
            'Stockage des abs d�but et fin de la section de travail du parcours
            unY1 = maColRepere(ComboRepDebSec.ListIndex + 1).monAbsCurv
            unY2 = maColRepere(ComboRepFinSec.ListIndex + 1).monAbsCurv
        End If
        ActualiserParcoursMoyen maColParcours(1), maColParcours, unY1, unY2
    End If
End Sub

Private Sub BtnMeteo_Click()
    frmModifMeteo.Show vbModal
End Sub

Private Sub BtnSuppPar_Click()
    Dim unPar As Parcours, unParUtil As Boolean
    Dim unY1 As Long, unY2 As Long
    Dim unNumPar As Long, unNbParUtil As Integer
    
    'Test pr�liminaire avant la destruction du feu
    unNumPar = SpreadParcours.ActiveRow
    If unNumPar = 1 Then
        MsgBox "La suppression du parcours moyen n'est pas autoris�e.", vbExclamation
    ElseIf unNumPar > 0 Then
        'R�cup�ration du parcours s�lectionn� et de son utilisation
        Set unPar = maColParcours(unNumPar)
        unParUtil = unPar.monIsUtil
        'Comptage du nombre de parcours utilis�
        unNbParUtil = DonnerNbParcoursUtil(Me)
        If (unNbParUtil = 1 And unParUtil) Or (unNbParUtil = 2 And maColParcours(1).monIsUtil) Then
            'Cas o� l'on veut supprimer le seul parcours utilis� autre
            'que le parcours moyen
            MsgBox "La suppression du seul parcours utilis� autre que le parcours moyen n'est pas autoris�e.", vbExclamation
        'If maColParcours.Count = 2 Then
            'Cas o� l'itin�raire ne contient plus qu'un parcours et donc aussi
            'le parcours moyen
            'MsgBox "La suppression du dernier parcours autre que le parcours moyen n'est pas autoris�e.", vbExclamation
        Else
            unMsg = "Etes-vous s�r de vouloir supprimer le parcours de la ligne n�"
            unMsg = unMsg + Format(unNumPar) + " nomm� " + unPar.monNom
            If MsgBox(unMsg, vbYesNo + vbQuestion) = vbYes Then
                'Cas de confirmation positive
                'D�selection dans les courbes DT et/ou DV si c'�tait
                'le parcours s�lectionn�
                If monIndParcoursSelectDT = unNumPar Then monIndParcoursSelectDT = 0
                If monIndParcoursSelectDV = unNumPar Then monIndParcoursSelectDV = 0
                'Suppression dans la collection des parcours
                maColParcours.Remove unNumPar
                'Suppression de la ligne du spread contenant le parcours
                SpreadParcours.Row = unNumPar
                SpreadParcours.Action = 5 ' = SS_ACTION_DELETE_ROW
                'Suppression de la ligne blanche ajout� en fin de spread
                SpreadParcours.MaxRows = SpreadParcours.MaxRows - 1
                'Recalcul du parcours moyen, de l'englobant et indication qu'il
                'faut tout redessiner si le parcours �tait utilis�
                If unParUtil Then
                    'Affichage pour calcul du parcours moyen
                    If CheckSection.Value = 0 Then
                        'Stockage des abs d�but et fin du parcours
                        unY1 = -100
                        unY2 = 1000000
                    Else
                        'Stockage des abs d�but et fin de la section de travail du parcours
                        unY1 = maColRepere(ComboRepDebSec.ListIndex + 1).monAbsCurv
                        unY2 = maColRepere(ComboRepFinSec.ListIndex + 1).monAbsCurv
                    End If
                    ActualiserParcoursMoyen maColParcours(1), maColParcours, unY1, unY2
                    'Calcul des englobants en temps et vitesse
                    CalculerEnglobantTV unY1, unY2
                    'Initialisation des indicateurs de redessin des onglets de 1 � 6
                    '� vrai pour d�clencher le dessin lors de leur activation
                    IndiquerToutRedessiner Me
                    'Mise � jour de la ligne 1 du spread parcours
                    'Celle contenant les info du parcours moyen
                    RemplirSpreadParcours Me, True
                End If
            End If
        End If
    Else
        MsgBox "Vous devez s�lectionner la ligne ou une des cellules de donn�es du parcours � supprimer.", vbInformation
    End If
End Sub

Private Sub BtnSuppRep_Click()
    SupprimerRepere Me, SpreadRepere.ActiveRow
End Sub

Private Sub CheckSection_Click()
    Dim uneSectionPasDefinie As Boolean
    Dim unY1 As Long, unY2 As Long
    Dim unPar As Parcours
      
    'Indication d'une modif
    maModif = True
    
    'Affichage du sablier en souris
    MousePointer = vbHourglass
                
    uneSectionPasDefinie = (CheckSection.Value = 0)
    ComboRepDebSec.Enabled = Not uneSectionPasDefinie
    ComboRepFinSec.Enabled = Not uneSectionPasDefinie
    BtnAjoutRep.Enabled = uneSectionPasDefinie
    BtnSuppRep.Enabled = uneSectionPasDefinie
    
    'On locke tout le spread spreadrepere en section de travail
    SpreadRepere.Col = -1
    SpreadRepere.Row = -1
    SpreadRepere.Lock = Not uneSectionPasDefinie
        
    If uneSectionPasDefinie Then
        'Stockage des abs d�but et fin du parcours
        unY1 = -100
        unY2 = 1000000
    Else
        'Stockage des abs d�but et fin de la section de travail du parcours
        unY1 = maColRepere(ComboRepDebSec.ListIndex + 1).monAbsCurv
        unY2 = maColRepere(ComboRepFinSec.ListIndex + 1).monAbsCurv
    End If
    
    'Calcul des englobants en temps et vitesse
    CalculerEnglobantTV unY1, unY2
    
    If uneSectionPasDefinie Then
        'On zoom en englobant tout l'itin�raire
        monMaxD = monMaxDtot
        monMinD = monMinDtot
        RedessinerZoomTout Me
    Else
        'on zoom en englobant toute la section de travail
        'entre debut et fin
        ZoomToutSection Me, unY1, unY2, ComboRepDebSec.ListIndex + 1
        'Suppression du message ci-dessous
        'MsgBox "Le tableau, contenant les informations des rep�res, n'est pas modifiable si une section de travail est d�finie.", vbInformation
    End If
    
    'Initialisation des min/max impression en distance, vitesse et temps
    'avec les valeurs des min/max automatique, juste la premi�re fois
    'et remise de la modif d'�chelle en impression � un nombre paire
    '==> pas de modif d'�chelle en impression
    maModifEchImp = 0
    monMinDImp = monMinD
    monMaxDImp = monMaxD
    monMinTImp = monMinT
    monMaxTImp = monMaxT
    monMinVImp = monMinV
    monMaxVImp = monMaxV
    
    'Restauration pointeur souris
    MousePointer = vbDefault
End Sub

Private Sub ComboRepDebSec_Click()
    Dim unY1 As Long, unY2 As Long
    
    'Cas o� la section de travail n'est pas coch�e, on ne fait rien
    If CheckSection.Value = 0 Then Exit Sub
    'Cas o� le rep�re de d�but de section de travail est le m�me, on ne fait rien
    If ComboRepDebSec.Text = ComboRepDebSec.Tag Then Exit Sub
    
    If ComboRepDebSec.Text = ComboRepFinSec.Text Then
        MsgBox "Les rep�res d�but et fin de section doivent �tre diff�rents.", vbCritical
        'On remet la derni�re valeur valide
        ComboRepDebSec.Text = ComboRepDebSec.Tag
    Else
        'Indication d'une modif
        maModif = True
        'Stockage dans le tag de la derni�re valeur valide
        ComboRepDebSec.Tag = ComboRepDebSec.Text
        If CheckSection.Value = 1 Then
            'Si on est en section de travail
            'on zoom en englobant toute la section de travail
            'entre debut et fin
            unY1 = maColRepere(ComboRepDebSec.ListIndex + 1).monAbsCurv
            unY2 = maColRepere(ComboRepFinSec.ListIndex + 1).monAbsCurv
            CalculerEnglobantTV unY1, unY2
            ZoomToutSection Me, unY1, unY2, ComboRepDebSec.ListIndex + 1
        End If
    End If
End Sub

Private Sub ComboRepFinSec_Click()
    Dim unY1 As Long, unY2 As Long
    
    'Cas o� la section de travail n'est pas coch�e, on ne fait rien
    If CheckSection.Value = 0 Then Exit Sub
    'Cas o� le rep�re de fin de section de travail est le m�me, on ne fait rien
    If ComboRepFinSec.Text = ComboRepFinSec.Tag Then Exit Sub
        
    If ComboRepDebSec.Text = ComboRepFinSec.Text Then
        MsgBox "Les rep�res d�but et fin de section doivent �tre diff�rents.", vbCritical
        'On remet la derni�re valeur valide
        ComboRepFinSec.Text = ComboRepFinSec.Tag
    Else
        'Indication d'une modif
        maModif = True
        'Stockage dans le tag de la derni�re valeur valide
        ComboRepFinSec.Tag = ComboRepFinSec.Text
        If CheckSection.Value = 1 Then
            'Si on est en section de travail
            'on zoom en englobant toute la section de travail
            'entre debut et fin
            unY1 = maColRepere(ComboRepDebSec.ListIndex + 1).monAbsCurv
            unY2 = maColRepere(ComboRepFinSec.ListIndex + 1).monAbsCurv
            CalculerEnglobantTV unY1, unY2
            ZoomToutSection Me, unY1, unY2, ComboRepFinSec.ListIndex + 1
        End If
    End If
End Sub



Private Sub Form_Activate()
    'Affectation de l'itin�raire courant
    DoEvents
    Set monIti = Me
    
    'R�organisation de la fen�tre fille si hauteur a boug�
    If PictureItiRef.Height <> ScaleHeight - 30 Then
        Form_Resize
    End If
    'Contexte d'aide de l'onglet
    frmMain.HelpContextID = HelpContextID
    Me.Show
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Then
        'Interdiction de taper des " sinon probl�mes de d�codages des chaines
        'dans les fichiers
        MsgBox "Les guillemets sont interdits, utilisez un autre caract�re.", vbInformation
        'Effacement du guillemet tap�
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    'Affectation de l'itin�raire courant
    DoEvents
    Set monIti = Me
    'Par d�faut, par de changement d'�chelle initialement � la cr�ation
    'd'une fen�tre d'itin�raire
    maModifEchImp = 0
    
    'Mise de la l�gende des noms de parcours horizontalement
    MSChart1.Legend.Location.LocationType = VtChLocationTypeBottomLeft
    'Affichage des classes de vitesses sur l'axe des x de l'histogramme
    With MSChart1.DataGrid
        .RowLabel(1, 1) = "[0, " + Format(mesOptions.maValClasV1) + "]"
        .RowLabel(2, 1) = "]" + Format(mesOptions.maValClasV1) + ", " + Format(mesOptions.maValClasV2) + "]"
        .RowLabel(3, 1) = "]" + Format(mesOptions.maValClasV2) + ", " + Format(mesOptions.maValClasV3) + "]"
        .RowLabel(4, 1) = "]" + Format(mesOptions.maValClasV3) + ", " + Format(mesOptions.maValClasV4) + "]"
        .RowLabel(5, 1) = "]" + Format(mesOptions.maValClasV4) + ", " + Format(mesOptions.maValClasV5) + "]"
        .RowLabel(6, 1) = "]" + Format(mesOptions.maValClasV5) + ", +oo["
    End With
    
    maModif = False 'Par d�faut pas de modif � l'ouverture
    
    'Par d�faut le premier rep�re est s�lectionn�
    monIndIconeRepSel = 1
    
    'Affichage �ventuelle de la section de travail
    ComboRepDebSec.Enabled = (CheckSection.Value = 1)
    ComboRepFinSec.Enabled = (CheckSection.Value = 1)
    
    'Taille des fen�tres filles
    Width = 0.95 * Screen.Width + 200 '120
    Height = 0.75 * Screen.Height
        
    'Retaillage de la picture box de visu des rep�res
    PictureItiRef.Top = 10
    PictureItiRef.Left = Width - PictureItiRef.Width - 90
    PictureItiRef.Height = ScaleHeight - 30
    AxeDist.Y1 = PictureItiRef.Height - 120
    
    'Retaillage des onglets
    TabData.Height = PictureItiRef.Height
    TabData.Top = PictureItiRef.Top
    TabData.Left = 0
    TabData.Width = PictureItiRef.Left - 10
    
    'Mise � jour des boutons dans la toolbar permettant l'impression
    'et la sauvegarde car il y a une fen�tre fille ouverte
    '==> Impression et sauvegarde possible
    frmMain.tbToolBar.Buttons("Print").Visible = True
    frmMain.tbToolBar.Buttons("Save").Visible = True
    
    'Mise � jour des items du menu Itin�raire permettant l'impression, la fermeture
    'et la sauvegarde (save et saveas) car il y a une fen�tre fille ouverte
    '==> Impression, fermeture et sauvegarde possible
    frmMain.mnuFileSave.Enabled = True
    frmMain.mnuFileSaveAs.Enabled = True
    frmMain.mnuFilePrint.Enabled = True
    frmMain.mnuFileClose.Enabled = True
    
    'Choix de la couleur des cellules lock�es dans les spreads Parcours et Repere
    'couleur des info-bulles souvent jaune
    SpreadParcours.LockBackColor = vbInfoBackground
    SpreadRepere.LockBackColor = vbInfoBackground
    'pour ces spreads tjs lock�es on prend juste la couleur de fond
    SpreadInfoParcoursDT.BackColor = vbInfoBackground
    SpreadInfoParcoursDV.BackColor = vbInfoBackground
    'Initialisation de la colonne du coef d'�talonnage des spreads
    'et les vitesses moy, min et max avec le s�parateur d�cimal en cours
    TrouverCaract�reD�cimalUtilis�
    InitColSpreadCaract�reD�cimal SpreadParcours, SpreadParcours.MaxCols, monCarDeci
    InitRowSpreadCaract�reD�cimal SpreadInfoParcoursDT, 5, monCarDeci
    InitRowSpreadCaract�reD�cimal SpreadInfoParcoursDT, 6, monCarDeci
    InitRowSpreadCaract�reD�cimal SpreadInfoParcoursDT, 7, monCarDeci
    InitRowSpreadCaract�reD�cimal SpreadInfoParcoursDV, 5, monCarDeci
    InitRowSpreadCaract�reD�cimal SpreadInfoParcoursDV, 6, monCarDeci
    InitRowSpreadCaract�reD�cimal SpreadInfoParcoursDV, 7, monCarDeci
    'Initialisation des l�gendes des classes de vitesses et de leur couleur
    'dans l'onglet synoptique des vitesses
    LabelClassV1.Caption = "0 <= V <= " + Format(mesOptions.maValClasV1)
    LabelClassV1.ForeColor = mesOptions.maCouleurClasV1
    LabelClassV2.Caption = Format(mesOptions.maValClasV1) + " < V <= " + Format(mesOptions.maValClasV2)
    LabelClassV2.ForeColor = mesOptions.maCouleurClasV2
    LabelClassV3.Caption = Format(mesOptions.maValClasV2) + " < V <= " + Format(mesOptions.maValClasV3)
    LabelClassV3.ForeColor = mesOptions.maCouleurClasV3
    LabelClassV4.Caption = Format(mesOptions.maValClasV3) + " < V <= " + Format(mesOptions.maValClasV4)
    LabelClassV4.ForeColor = mesOptions.maCouleurClasV4
    LabelClassV5.Caption = Format(mesOptions.maValClasV4) + " < V <= " + Format(mesOptions.maValClasV5)
    LabelClassV5.ForeColor = mesOptions.maCouleurClasV5
    LabelClassV6.Caption = "V > " + Format(mesOptions.maValClasV5)
    LabelClassV6.ForeColor = mesOptions.maCouleurClasV6
    
    'Affichage des libell�s de la ligne d'ent�te du tableau brut
    SpreadTabBrut.Row = 0
    
    SpreadTabBrut.Col = 1
    unTitreCol = "Information Tron�on : " + Chr(13) + "Rep�re D�but"
    unTitreCol = unTitreCol + Chr(13) + "Abscisse D�but (m)"
    unTitreCol = unTitreCol + Chr(13) + "Rep�re Fin"
    unTitreCol = unTitreCol + Chr(13) + "Abscisse Fin (m)"
    SpreadTabBrut.Text = unTitreCol
    
    SpreadTabBrut.Col = 2
    SpreadTabBrut.Text = "Parcours" + Chr(13) + "Jour, Date et Heure de mesure"
    
    SpreadTabBrut.Col = 3
    SpreadTabBrut.Text = "Distance et" + Chr(13) + "Temps de" + Chr(13) + "parcours"
    
    SpreadTabBrut.Col = 5
    unTitreCol = "Dur�e Arr�ts" + Chr(13) + "Nombre Arr�ts"
    unTitreCol = unTitreCol + Chr(13) + "% Tps Arr�ts"
    SpreadTabBrut.Text = unTitreCol
    
    SpreadTabBrut.Col = 6
    unTitreCol = "Dur�e Dbl Top" + Chr(13) + "Nbre Dbl Top"
    unTitreCol = unTitreCol + Chr(13) + "% Tps Dbl Top"
    SpreadTabBrut.Text = unTitreCol
    
    'Affichage des libell�s de la ligne d'ent�te du tableau Synth�se/Statistique
    SpreadTabSS.Row = 0
    
    SpreadTabSS.Col = 1
    unTitreCol = "Information Tron�on : " + Chr(13) + "Rep�re D�but"
    unTitreCol = unTitreCol + Chr(13) + "Abscisse D�but (m)"
    unTitreCol = unTitreCol + Chr(13) + "Rep�re Fin"
    unTitreCol = unTitreCol + Chr(13) + "Abscisse Fin (m)"
    unTitreCol = unTitreCol + Chr(13) + "Longueur (m)"
    SpreadTabSS.Text = unTitreCol
        
    'Pour remettre les fen�tres bien arrang�es en cascade
    'D�calage des fen�tres de 330 twpis en X et Y � chaque fen�tre
    Top = (Forms.Count - 2) * 330
    Left = (Forms.Count - 2) * 330
    frmMain.Arrange vbCascade
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim unNomFich As String
    
    'Indication d'une modif
    SpreadRepere.EditMode = False
    SpreadParcours.EditMode = False
    DoEvents 'Sortie du mode saisie �ventuelle et vidage des events
             'pour que les changemade suivant soit synchro
    If SpreadRepere.ChangeMade Then maModif = True
    If SpreadParcours.ChangeMade Then maModif = True
    
    'Demande de sauvegarde si modif
    If maModif Or EstNouvelIti(Me) Then
        If EstNouvelIti(Me) Then
            unNomFich = Caption
        Else
            unNomFich = Mid(Caption, 12)
        End If
        uneRep = MsgBox(MsgSaveFile + unNomFich + " ?", vbExclamation + vbYesNoCancel)
        If uneRep = vbCancel Then
            'Pas de sortie, on ne fait rien
            Cancel = True
        ElseIf uneRep = vbYes Then
            'Sauvegarde puis sortie
            If SauverFichier(Me, unNomFich, False) = "" Then
                'Si pas de fichier choisie ==> on ne sort pas
                Cancel = True
            Else
                Cancel = False
            End If
        Else
            'Cas du click sur Non ==> On sort
            Cancel = False
        End If
    End If
End Sub

Private Sub Form_Resize()
    'L'event resize ne se produit qu'au lancement (form_load)
    'et lors d'une mise en plein �cran (par la case carr� en haut
    '� droite dans le titre entre le _ et le x ou par l'item
    'Agrandissement du menu d�roulant de fen�tre fille obtenu par
    'click gauche sur l'ic�ne en haut � gauche)
    Dim uneMargeG As Single, uneMargeD As Single
    Dim uneMargeH As Single, uneMargeB As Single, unMaxYecran As Single
    
    'Si la fen�tre n'est pas encore affich�e ou mise en ic�ne
    '==> on ne fait rien
    If maColRepere Is Nothing Or Me.Visible = False Or Me.WindowState = vbMinimized Then Exit Sub
    
    frmMain.MousePointer = vbHourglass
            
    'Initialisation des indicateurs de redessin des onglets de 1 � 6
    '� vrai pour d�clencher le dessin lors de leur activation
    IndiquerToutRedessiner Me
    
    'Placement de la frame de visu de l'itin�raire de r�f�rence
    PictureItiRef.Height = ScaleHeight - 30
    PictureItiRef.Left = Width - PictureItiRef.Width - 90
    AxeDist.Y1 = PictureItiRef.Height - 120
    
    'Retaillage des onglets
    TabData.Width = PictureItiRef.Left - 10
    TabData.Height = PictureItiRef.Height
    
    'Mise au premier plan des onglets
    TabData.ZOrder 0
    
    'Retaillage de l'onglet courant
    Select Case TabData.Tab
    Case OngletItiRef
        RetaillerOngletItiRef Me
    Case OngletCbeDT
        RetaillerOngletCbeDT Me
    Case OngletCbeDV
        RetaillerOngletCbeDV Me
    Case OngletSynoV
        RetaillerOngletSynoV Me
    Case OngletHistV
        RetaillerOngletHistV Me
    Case OngletTabBr
        RetaillerOngletTabBr Me
    Case OngletTabSS
        RetaillerOngletTabSS Me
    Case Else   ' Autres valeurs.
        MsgBox MsgErreurProg + MsgErreurNumOngletInconnu + MsgIn + "frmDocument:Resize", vbCritical
    End Select
    
    'Fixer les tailles de marges et la longueur �cran en y maxi
    'FixerMargesPicBox Me, Me.PicBoxDT, uneMargeG, uneMargeD, uneMargeH, uneMargeB
    'Redessin des rep�res au bon zoom
    PictureItiRef.Cls
    For i = 1 To maColRepere.Count
        DessinerRepere Me, maColRepere(i)
    Next i
        
    frmMain.MousePointer = vbDefault
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ViderColParcours maColParcours
    ViderCollection maColMeteo
    ViderColRepere maColRepere
    
    If Forms.Count = 2 Then
        'Fermeture de la seule fen�tre fille, il reste deux fen�tres
        'la MDI m�re et la seule fille pas encore d�truite
        'Mise � jour des boutons dans la toolbar permettant l'impression
        'et la sauvegarde car il n'y a pas de fen�tre fille ouverte
        '==> Impression et sauvegarde impossible
        frmMain.tbToolBar.Buttons("Print").Visible = False
        frmMain.tbToolBar.Buttons("Save").Visible = False
    
        'Mise � jour des items du menu Itin�raire permettant l'impression, la fermeture
        'et la sauvegarde (save et saveas) car il n'y a pas de fen�tre fille ouverte
        '==> Impression, fermeture et sauvegarde impossible
        frmMain.mnuFileSave.Enabled = False
        frmMain.mnuFileSaveAs.Enabled = False
        frmMain.mnuFilePrint.Enabled = False
        frmMain.mnuFileClose.Enabled = False
    
        'Remise � z�ro du contexte d'aide pour ouvrir sur le sommaire lors du F1
        frmMain.HelpContextID = 0
    End If
    
    'Fermeture du fichier ==> utilisable par une autre appli (MiTemps ou �diteur)
    Close #monFichId
End Sub






Private Sub ImageRepere_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Affichage de la bordure pour montrer la s�lection
    Dim unRep As Repere, uneRow As Integer
    
    'Stockage dans le tag de la fen�tre fille de la cl� d'identification
    'correspond � ce rep�re
    Tag = ImageRepere(Index).Tag
    
    DeselectionnerRepere Me, SpreadRepere.ActiveRow
    uneRow = DonnerLigneRepere
    SelectionnerRepere Me, uneRow
End Sub

Private Sub ImageRepere_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Traitement des click
    If Button = vbKeyRButton Then
        frmMain.AfficherMenuContextuel Index
    End If
End Sub

Private Sub PicBoxDT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'S�lection d'un parcours et �paisseur de la courbe
    'mis en plus gros
    unIndParcoursSelectDT = SelectionnerParcours(Me, X, Y)
    If unIndParcoursSelectDT = 0 Then
        If monIndParcoursSelectDT > 0 Then
            unMsg = "Aucun parcours s�lectionn�, Redessin des courbes Distance/Temps avec la m�me �paisseur."
        Else
            unMsg = "Aucun parcours s�lectionn�."
        End If
        MsgBox unMsg, vbInformation
    End If
    If unIndParcoursSelectDT <> monIndParcoursSelectDT Then
        'Si on clique un autre, on redessine les courbes DT
        'la fonction DessinerCourbe mettra en trait gros la s�lection
        monIndParcoursSelectDT = unIndParcoursSelectDT
        If PicBoxDT.Tag = "" Then
            'On ne redessine que si le parcours n'a pas �t� choisi
            'dans une liste de plusieurs (fen�tre frmChoixPar)
            DessinerCourbe Me, PicBoxDT, OngletCbeDT
        End If
    End If
End Sub

Private Sub PicBoxDV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'S�lection d'un parcours et �paisseur de la courbe
    'mis en plus gros
    unIndParcoursSelectDV = SelectionnerParcours(Me, X, Y)
    If unIndParcoursSelectDV = 0 Then
        If monIndParcoursSelectDV > 0 Then
            unMsg = "Aucun parcours s�lectionn�, Redessin des courbes Distance/Vitesse avec la m�me �paisseur."
        Else
            unMsg = "Aucun parcours s�lectionn�."
        End If
        MsgBox unMsg, vbInformation
    End If
    If unIndParcoursSelectDV <> monIndParcoursSelectDV Then
        'Si on clique un autre, on redessine les courbes DT
        'la fonction DessinerCourbe mettra en trait gros la s�lection
        monIndParcoursSelectDV = unIndParcoursSelectDV
        If PicBoxDV.Tag = "" Then
            'On ne redessine que si le parcours n'a pas �t� choisi
            'dans une liste de plusieurs (fen�tre frmChoixPar)
            DessinerCourbe Me, PicBoxDV, OngletCbeDV
        End If
    End If
End Sub




Private Sub SpreadParcours_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim unPar As Parcours, uneAbs1 As Long, uneAbs2 As Long
    
    If Col = 2 Then
        'Modification du champ utilisation du parcours de la ligne s�lectionn�e
        SpreadParcours.Row = Row
        SpreadParcours.Col = Col
        'Test qu'il ne reste pas que le parcours moyen en utilisation
        unNbParUtil = DonnerNbParcoursUtil(Me)
        If unNbParUtil = 0 And Row = 1 And SpreadParcours.Value = 1 Then
            MsgBox "S�lection impossible, car il n'aurait plus de parcours s�lectionn� � part le parcours moyen, donc impossible de calculer les moyennes.", vbExclamation
            'On des�lectionne
            SpreadParcours.Value = 0
            Exit Sub
        ElseIf SpreadParcours.Visible And unNbParUtil = 2 And Row <> 1 And SpreadParcours.Value = 0 And maColParcours(1).monIsUtil Then
            MsgBox "D�selection impossible, car il n'aurait plus de parcours s�lectionn� � part le parcours moyen, donc impossible de calculer les moyennes.", vbExclamation
            'On res�lectionne
            SpreadParcours.Value = 1
            Exit Sub
        End If
        'pour les checkbox, value = 1 si coch�e, 0 si d�coch�e
        maColParcours(Row).monIsUtil = (SpreadParcours.Value = 1)
        Set unPar = maColParcours(1)
        If CheckSection.Value = 0 Then
            'Cas o� on travaille sur tout le parcours
            'Stockage des abs d�but et fin du parcours
            uneAbs1 = -100
            uneAbs2 = 1000000
        Else
            'Cas o� on travaille sur une section du parcours
            'Stockage des abs d�but et fin de la section de travail du parcours
            uneAbs1 = maColRepere(ComboRepDebSec.ListIndex + 1).monAbsCurv
            uneAbs2 = maColRepere(ComboRepFinSec.ListIndex + 1).monAbsCurv
        End If
        If SpreadParcours.Visible Then
            'Le calcul d'actualisation du parcours moyen ne se fait que si
            'le spreadparcours est visible, sinon � l'ouverture d'un
            'fichier itin�raire (= *.mit) ou � la cr�ation d'un itin�raire �
            'partir d'un fichier mtb, ce calcul est fait � chaque fois que l'on
            'fait une affection des cases � cocher du spreadparcours
            ActualiserParcoursMoyen unPar, maColParcours, uneAbs1, uneAbs2
            'Calcul des englobants en temps et vitesse
            CalculerEnglobantTV uneAbs1, uneAbs2
            'Initialisation des indicateurs de redessin des onglets de 1 � 6
            '� vrai pour d�clencher le dessin lors de leur activation
            IndiquerToutRedessiner Me
            'Mise � jour de la ligne 1 du spread parcours
            'Celle contenant les info du parcours moyen
            RemplirSpreadParcours Me, True
        End If
    End If
End Sub


Private Sub SpreadParcours_Change(ByVal Col As Long, ByVal Row As Long)
    'Stockage des changements dans les colonnes 1 � 9 sauf 3 et 4
    'qui sont fait dans les events Click et ButtonClicked de spreadparcours
    Dim unPar As Parcours
    
    'R�cup du parcours
    Set unPar = maColParcours(Row)
    'Positionnement sur la cellule active
    SpreadParcours.Row = Row
    SpreadParcours.Col = Col
    
    If Col = 1 Then
        'Stockage du nom du parcours
        unPar.monNom = SpreadParcours.Text
    ElseIf Col = 4 Then
        'Stockage du nom de l'enqu�teur
        unPar.monEnqueteur = SpreadParcours.Text
    ElseIf Col = 5 Then
        'Stockage du nom de l'enqu�teur
        unPar.monNumVeh = SpreadParcours.Text
    ElseIf Col = 6 Then
        'Cas du changement par click dans la liste de la combobox
        'des conditions m�t�o
        unPar.maMeteo = Format(Mid(SpreadParcours.Text, 1, 1))
    ElseIf Col = 7 Then
        'Stockage de la date de mesure
        unPar.maDate = SpreadParcours.Text
        unJour = DonnerJourSemaine(CDate(SpreadParcours.Text))
        SpreadParcours.Col = 8
        'Mise � jour du jour de semaine
        SpreadParcours.Text = unJour
        'Stockage du jour de mesure
        unPar.monJourSemaine = SpreadParcours.Text
        'Indication du redessin de l'onglet Tableau Brut
        SetTabRedOnglet OngletTabBr, True
    ElseIf Col = 8 Then
        'Test de la coh�rence entre la date et le jour de la semaine
        SpreadParcours.Col = Col 'colonne jour mesure
        unJourMesure = SpreadParcours.Text
        SpreadParcours.Col = 7 'colonne date mesure
        unJour = DonnerJourSemaine(CDate(SpreadParcours.Text))
        If unJourMesure <> unJour Then
            MsgBox "Le " + SpreadParcours.Text + " n'est pas un " + unJourMesure + " mais un " + unJour, vbExclamation
            SpreadParcours.Col = 8 'colonne jour mesure
            SpreadParcours.Text = unJour
        Else
            'Indication du redessin de l'onglet Tableau Brut
            SetTabRedOnglet OngletTabBr, True
        End If
        'Stockage du jour de mesure
        unPar.monJourSemaine = SpreadParcours.Text
    End If
End Sub

Private Sub SpreadParcours_DblClick(ByVal Col As Long, ByVal Row As Long)
'Private Sub SpreadParcours_Click(ByVal Col As Long, ByVal Row As Long)
    If Col = 3 Then
        'Ouverture de la fen�tre de changement de couleur
        ' Attribue � CancelError la valeur True
        frmMain.dlgCommonDialog.CancelError = True
        On Error GoTo ErrColorCancel
        ' D�finit la propri�t� Flags
        frmMain.dlgCommonDialog.flags = cdlCCRGBInit
        ' Affiche la bo�te de dialogue Couleur
        frmMain.dlgCommonDialog.ShowColor
        ' Attribue � l'arri�re-plan de la feuille la
        ' couleur s�lectionn�e
        SpreadParcours.Col = Col
        SpreadParcours.Row = Row
        SpreadParcours.BackColor = frmMain.dlgCommonDialog.Color
        maColParcours(Row).maCouleur = frmMain.dlgCommonDialog.Color
        'Indication d'une modif
        maModif = True
        'Initialisation des indicateurs de redessin des onglets de 1 � 4
        '� vrai pour d�clencher le dessin lors de leur activation
        'Tous les onglet sauf ItiRef, TabBrut et TabStat
        IndiquerToutRedessiner Me, OngletCbeDT, OngletHistV
    End If
        
    On Error GoTo 0
    Exit Sub

ErrColorCancel:
    ' L'utilisateur a cliqu� sur Annuler
    On Error GoTo 0
End Sub


Private Sub SpreadParcours_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim unJour As String, unJourMesure As String, unPar As Parcours
    
    'R�cup du parcours
    Set unPar = maColParcours(SpreadParcours.ActiveRow)
    SpreadParcours.Row = SpreadParcours.ActiveRow
    SpreadParcours.Col = SpreadParcours.ActiveCol
    
    If SpreadParcours.ActiveCol = 1 Then
        'Stockage du nom du parcours
        unPar.monNom = SpreadParcours.Text
        'Initialisation des indicateurs de redessin des onglets de 3 � 5
        '� vrai pour d�clencher le dessin lors de leur activation
        'Redessin des onglets : SynoV, HistoV, TabBrut
        IndiquerToutRedessiner Me, OngletSynoV, OngletTabBr
    ElseIf SpreadParcours.ActiveCol = 4 Then
        'Stockage du nom de l'enqu�teur
        unPar.monEnqueteur = SpreadParcours.Text
    ElseIf SpreadParcours.ActiveCol = 5 Then
        'Stockage du nom de l'enqu�teur
        unPar.monNumVeh = SpreadParcours.Text
    ElseIf SpreadParcours.ActiveCol = 6 Then
        'Cas du changement par click dans la liste de la combobox
        'des conditions m�t�o
        unPar.maMeteo = Format(Mid(SpreadParcours.Text, 1, 1))
    ElseIf SpreadParcours.ActiveCol = 7 Then
        'Mise en coh�rence entre la date et le jour de la semaine
        'Stockage de la date de mesure
        unPar.maDate = SpreadParcours.Text
        unJour = DonnerJourSemaine(CDate(SpreadParcours.Text))
        SpreadParcours.Col = 8
        'Mise � jour du jour de semaine
        SpreadParcours.Text = unJour
        'Stockage du jour de mesure
        unPar.monJourSemaine = SpreadParcours.Text
    ElseIf SpreadParcours.ActiveCol = 8 Then
        'Test de la coh�rence entre la date et le jour de la semaine
        SpreadParcours.Row = SpreadParcours.ActiveRow
        SpreadParcours.Col = 8 'colonne jour mesure
        unJourMesure = SpreadParcours.Text
        SpreadParcours.Col = 7 'colonne date mesure
        unJour = DonnerJourSemaine(CDate(SpreadParcours.Text))
        If unJourMesure <> unJour Then
            'Correction du jour de la semaine
            SpreadParcours.Col = 8 'colonne jour mesure
            SpreadParcours.Text = unJour
            'MsgBox "Le " + SpreadParcours.Text + " n'est pas un " + unJourMesure + " mais un " + unJour, vbExclamation
        End If
    ElseIf SpreadParcours.ActiveCol = 9 Then
        'Stockage de l'heure de mesure
        unPar.monHeureDebut = SpreadParcours.Text
        'Indication du redessin de l'onglet Tableau Brut
        SetTabRedOnglet OngletTabBr, True
    End If
End Sub


Private Sub SpreadParcours_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    'Pour corriger un bug spread en vb quand le scroller bouge
    DoEvents
    SpreadRepere.Refresh
End Sub

Private Sub SpreadRepere_Change(ByVal Col As Long, ByVal Row As Long)
    Dim unRep As Repere, uneCle As String
    
    'R�cup�ration du rep�re en cours de modif
    'sa cl� est dans la derni�re colonne de la ligne active
    SpreadRepere.Row = SpreadRepere.ActiveRow
    SpreadRepere.Col = SpreadRepere.MaxCols
    uneCle = SpreadRepere.Text
    Set unRep = maColRepere(uneCle)

    'Traitement de la modif
    SpreadRepere.Row = SpreadRepere.ActiveRow
    SpreadRepere.Col = SpreadRepere.ActiveCol
    If SpreadRepere.ActiveCol = 5 Then
        'Modif de l'ic�ne du rep�re dans les colonnes 4 et 5 et sur l'axe des distances
        ModifierIconeRepere Me, unRep
        'Indication de modif
        maModif = True
    End If
End Sub

Private Sub SpreadRepere_KeyPress(KeyAscii As Integer)
    'Stockage de l'ancienne valeur pour remise � l'�tat initial du nom court
    'en cas d'apparition de la fen�tre de saisie pour avoir unicit� du nom court
    'et de click sur le bouton annuler de cette fen�tre
    SpreadRepere.Row = SpreadRepere.ActiveRow
    SpreadRepere.Col = SpreadRepere.ActiveCol
    SpreadRepere.Tag = SpreadRepere.Text
End Sub

Private Sub SpreadRepere_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim unRep As Repere, uneCle As String, uneAbs As Long
    Dim unNomCourt As String, uneContinuationBoucle As Boolean
    
    'Indication d'une modif
    If SpreadRepere.ChangeMade Then maModif = True
    
    'Initialisation
    uneContinuationBoucle = True
    
    'R�cup�ration du rep�re en cours de modif
    'sa cl� est dans la derni�re colonne de la ligne active
    SpreadRepere.Row = SpreadRepere.ActiveRow
    SpreadRepere.Col = SpreadRepere.MaxCols
    uneCle = SpreadRepere.Text
    Set unRep = maColRepere(uneCle)
    
    'Traitement de la modif
    SpreadRepere.Row = SpreadRepere.ActiveRow
    SpreadRepere.Col = SpreadRepere.ActiveCol
    If SpreadRepere.ActiveCol = 1 Then
        'Modif du nom long du rep�re
        unRep.monNomLong = SpreadRepere.Text
    ElseIf SpreadRepere.ActiveCol = 2 Then
        'Modif du nom court du rep�re avec v�rif d'unicit�
        'V�rification de l'unicit� du nom court
        If unRep.monNomCourt <> SpreadRepere.Text Then
            'Cas o� une touche modifiant le nom a �t� tap�
            'On r�gle ainsi les fl�ches, les suppr...
            unNomCourt = SpreadRepere.Text
            While uneContinuationBoucle
                uneContinuationBoucle = VerifierNomCourtUnique(Me, unNomCourt) = False Or Len(unNomCourt) > 15
                If uneContinuationBoucle Then
                    'Demande de saisie d'un autre nom
                    unNomCourt = InputBox("Un rep�re poss�de d�j� le nom court ci-dessous, entrez un nouveau nom court (15 caract�res maximum).", , unNomCourt)
                    If unNomCourt = "" Then
                        unNomCourt = SpreadRepere.Tag
                        uneContinuationBoucle = False
                    ElseIf InStr(1, unNomCourt, Chr(34)) > 0 Then
                        'Interdiction de taper des " sinon probl�mes de d�codages des chaines
                        'dans les fichiers
                        MsgBox "Les guillemets sont interdits, utilisez un autre caract�re.", vbInformation
                        unNomCourt = SpreadRepere.Text
                        uneContinuationBoucle = True
                    End If
                    SpreadRepere.Text = unNomCourt
                Else
                    'Sortie de la boucle
                    uneContinuationBoucle = False
                End If
            Wend
            unRep.monNomCourt = unNomCourt
            'Modif du libell� dans les combobox de d�but et fin de section
            'Ce libell� se trouve en position num ligne courant - 1
            ComboRepDebSec.List(SpreadRepere.ActiveRow - 1) = unRep.monNomCourt
            ComboRepFinSec.List(SpreadRepere.ActiveRow - 1) = unRep.monNomCourt
            'Modif de l'info-bulle
            unRep.monIcone.ToolTipText = unRep.monNomCourt + " / Type : " + DonnerIconeRepere(unRep.monTypeIcone).Tag + " / AbsCurv = " + Format(unRep.monAbsCurv) + " m"
            'Indication de redessin des onglets Tableau Brut et synth�se/Stats
            Me.SetTabRedOnglet OngletTabBr, True
            Me.SetTabRedOnglet OngletTabSS, True
        End If
    ElseIf SpreadRepere.ActiveCol = 3 Then
        'Modif de l'abscisse curviligne (= Distance) du rep�re
        uneAbs = CLng(SpreadRepere.Text)
        'Teste si abscisse curviligne unique
        i = 1
        While i <= maColRepere.Count
            If maColRepere(i).monAbsCurv = uneAbs And Not maColRepere(i) Is unRep Then
                unMsgInfo = "Il existe d�j� un rep�re � l'abscisse curviligne valant " + Format(uneAbs) + " m�tres"
                unMsgInfo = unMsgInfo + Chr(13) + Chr(13) + "Entrer une autre abscisse curviligne compris entre 0 et 1 000 000 m�tres"
                uneRepText = InputBox(unMsgInfo, App.Title + " : Abscisse curviligne existant", Format(uneAbs))
                If uneRepText = "" Then
                    'Si r�ponse vide ou click sur annuler
                    '==> remise de la valeur pr�c�dente valide
                    uneAbs = unRep.monAbsCurv
                    SpreadRepere.Text = Format(uneAbs)
                    i = maColRepere.Count 'sortie boucle While (= valeur limite)
                Else
                    i = 0
                    uneAbs = CLng(uneRepText)
                    SpreadRepere.Text = Format(uneAbs)
                End If
            End If
            'Incr�mentation pour le coup suivant
            i = i + 1
        Wend
        ModifierAbsCurvRepere Me, unRep, uneAbs
        'Indication de redessin des onglets Tableau Brut et synth�se/Stats
        'pour les courbes faits dans RedessinerZoomTout car le zoom change
        Me.SetTabRedOnglet OngletTabBr, True
        Me.SetTabRedOnglet OngletTabSS, True
    ElseIf SpreadRepere.ActiveCol = 5 Then
        'Modif de l'ic�ne du rep�re dans les colonnes 4 et 5 et sur l'axe des distances
        ModifierIconeRepere Me, unRep
    Else
        'Cas d'une erreur de programmation,
        'sauf si colonne 4 car on n'y fait rien
        If SpreadRepere.ActiveCol <> 4 Then
            MsgBox MsgErreurProg + MsgErreurNumColInconnu + MsgIn + "frmDocument:SpreadRepere_KeyUp", vbCritical
        End If
    End If
End Sub

Private Sub SpreadRepere_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow > -1 Then
        DeselectionnerRepere Me, CInt(Row)
        SelectionnerRepere Me, CInt(NewRow), CInt(NewCol)
    End If
    If SpreadRepere.ChangeMade Then maModif = True
End Sub



Private Sub SpreadRepere_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    'Pour corriger un bug spread en vb quand le scroller bouge
    DoEvents
    SpreadRepere.Refresh
End Sub

Private Sub TabData_Click(PreviousTab As Integer)
    'Retaillage de l'onglet courant
    DoEvents
    'Test s'il y a des parcours utilis�s
    If maColParcours.Count > 0 And DonnerNbParcoursUtil(Me) = 0 And TabData.Tab <> OngletItiRef And TabData.TabEnabled(TabData.Tab) = True Then
        MsgBox "Dans l'" + Caption + Chr(13) + Chr(13) + "Aucun r�sultat n'est disponible car aucun parcours n'est utilis� (cf colonne [Utilis�] du tableau des parcours de l'onglet Itin�raire de r�f�rence).", vbExclamation
        TabData.Tab = OngletItiRef
    End If
    Select Case TabData.Tab
        Case OngletItiRef
            RetaillerOngletItiRef Me
            'Contexte d'aide
            frmMain.HelpContextID = HelpID_WinOngletItiRef
        Case OngletCbeDT
            RetaillerOngletCbeDT Me, True
            'Remplissage des info sur le parcours s�lectionn� dans l'onglet courbe DT
            RemplirSpreadInfoParcoursSel SpreadInfoParcoursDT, Me, monIndParcoursSelectDT
            'Contexte d'aide
            frmMain.HelpContextID = HelpID_WinOngletCbeDT
            Me.PicBoxDT.SetFocus
        Case OngletCbeDV
            RetaillerOngletCbeDV Me, True
            'Remplissage des info sur le parcours s�lectionn� dans l'onglet courbe DV
            RemplirSpreadInfoParcoursSel SpreadInfoParcoursDV, Me, monIndParcoursSelectDV
            'Contexte d'aide
            frmMain.HelpContextID = HelpID_WinOngletCbeDV
            'HelpContextID = HelpID_WinOngletCbeDV
            Me.PicBoxDV.SetFocus
        Case OngletSynoV
            RetaillerOngletSynoV Me, True
            'Contexte d'aide
            frmMain.HelpContextID = HelpID_WinOngletSynoV
            Me.PicBoxSynoV.SetFocus
        Case OngletHistV
            RetaillerOngletHistV Me, True
            'Contexte d'aide
            frmMain.HelpContextID = HelpID_WinOngletHistoV
            If MSChart1.Visible Then MSChart1.SetFocus
        Case OngletTabBr
            RetaillerOngletTabBr Me, True
            'Contexte d'aide
            frmMain.HelpContextID = HelpID_WinOngletTabBrut
            Me.SpreadTabBrut.SetFocus
        Case OngletTabSS
            RetaillerOngletTabSS Me, True
            'Contexte d'aide
            frmMain.HelpContextID = HelpID_WinOngletTabStat
            Me.SpreadTabSS.SetFocus
        Case Else   ' Autres valeurs.
            MsgBox MsgErreurProg + MsgErreurNumOngletInconnu + MsgIn + "frmDocument:Resize", vbCritical
    End Select
    'Contexte d'aide de l'onglet
    HelpContextID = frmMain.HelpContextID
End Sub


Private Sub CalculerEnglobantTV(unY1 As Long, unY2 As Long)
    'Proc�dure calculant les valeurs mini et maxi des temps et vitesses
    'pour faire un zoom englobant correcte dans les courbes DT et DV
    'R�initialisation des temps et vitesses maxi et mini
    MousePointer = vbHourglass
    monMaxT = 0
    monMinT = 1000000
    monMaxV = 0
    monMinV = 0 '1000000 'pour avoir ls cas o� la mesure d�marrer
    'en cours de roulage
    
    'Mise � jour de la longueur de l'itin�raire
    TextLongIti.Text = Format(DonnerLongIti(Me))
    
    'Calcul des infos sur chaque parcours
    For i = 1 To maColParcours.Count
        Set unPar = maColParcours(i)
        If unPar.monIsUtil Then
            'Calcul des vitesses min, max et moyenne et de la dur�e, du nombre
            'et du temps d'arr�ts sur le parcours total
            'Conversion des abs curvilignes des m�tres en d�cim�tres
            '(d�cim�tre = unit� des pas de mesure) d'o� le * 10
            unPar.CalculerLesVitDistDureeEtArrets unY1 * 10, unY2 * 10
            If monMinV > unPar.maVmin Then monMinV = unPar.maVmin
            If monMaxV < unPar.maVmax Then monMaxV = unPar.maVmax
            If monMinT > unPar.monTDebSection Then monMinT = unPar.monTDebSection
            If monMaxT < unPar.monTFinSection Then monMaxT = unPar.monTFinSection
            'Calcul du nombre et du temps de double tops sur le parcours total
            'Conversion des abs curvilignes des m�tres en d�cim�tres
            '(d�cim�tre = unit� des pas de mesure) d'o� le * 10
            unPar.CalculerNbEtDureeDoubleTop unY1 * 10, unY2 * 10
        End If
    Next i
    
    'Conversion du temps mini et maxi des dixi�mes de secondes en minutes
    monMaxT = monMaxT / 600
    monMinT = monMinT / 600
    
    MousePointer = vbDefault
End Sub

Private Sub TextNomIti_Change()
    'Indication d'une modif
    maModif = True
    'Stockage du nom d'itin�raire
    monNomIti = TextNomIti.Text
End Sub


Public Function GetTabRedOnglet(unNum As Integer) As Boolean
    'Fonction donnant la valeur en position i du tableau monTabRedOnglet
    GetTabRedOnglet = monTabRedOnglet(unNum)
End Function

Public Sub SetTabRedOnglet(unNum As Integer, unBool As Boolean)
    'Fonction modifiant la valeur en position i du tableau monTabRedOnglet
    monTabRedOnglet(unNum) = unBool
End Sub


Public Sub AppelerCheckSectionClick()
    'Proc�dure permettant d'appeler de l'ext�rieur le contenu du click event
    'de la case � cocher CheckSection correspondant au section de travail
    'Cette fonction ne sert qu'une fois apr�s l'ouverture du fichier mit
    CheckSection_Click
    maModif = False 'Car checksection_click met maModif � true
End Sub
