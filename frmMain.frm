VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "MiTemps"
   ClientHeight    =   4695
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6675
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "Nouveau"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "Ouvrir"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Enregistrer"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimer"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1440
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4425
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6112
            Text            =   "État"
            TextSave        =   "État"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "23/02/2006"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11:20"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   960
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ListIcons 
      Left            =   2520
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   33
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
            Object.Tag             =   "Feu tricolore"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":075C
            Key             =   ""
            Object.Tag             =   "Panneau stop"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0A76
            Key             =   ""
            Object.Tag             =   "Cédez le passage"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0D90
            Key             =   ""
            Object.Tag             =   "Carrefour"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":10AA
            Key             =   ""
            Object.Tag             =   "Giratoire"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":13C4
            Key             =   ""
            Object.Tag             =   "Entrée d'agglomération"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":16DE
            Key             =   ""
            Object.Tag             =   "Sortie d'agglomération"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":19F8
            Key             =   ""
            Object.Tag             =   "Arrêt de bus"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1D12
            Key             =   ""
            Object.Tag             =   "Passage piéton"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":202C
            Key             =   ""
            Object.Tag             =   "Péage"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2346
            Key             =   ""
            Object.Tag             =   "Entrée d'autoroute"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2660
            Key             =   ""
            Object.Tag             =   "Sortie d'autoroute"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":297A
            Key             =   ""
            Object.Tag             =   "Station service"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2C94
            Key             =   ""
            Object.Tag             =   "Aire de repos"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2FAE
            Key             =   ""
            Object.Tag             =   "Autre"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":32C8
            Key             =   ""
            Object.Tag             =   "Double Top"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":35E2
            Key             =   ""
            Object.Tag             =   "Fin de toutes interdictions"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":38FC
            Key             =   ""
            Object.Tag             =   "Début de limitation de vitesse à 30 km/h"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3C16
            Key             =   ""
            Object.Tag             =   "Fin de limitation de vitesse à 30 km/h"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3F30
            Key             =   ""
            Object.Tag             =   "Début de limitation de vitesse à 50 km/h"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":424A
            Key             =   ""
            Object.Tag             =   "Fin de limitation de vitesse à 50 km/h"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4564
            Key             =   ""
            Object.Tag             =   "Début de limitation de vitesse à 70 km/h"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":487E
            Key             =   ""
            Object.Tag             =   "Fin de limitation de vitesse à 70 km/h"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4B98
            Key             =   ""
            Object.Tag             =   "Début de limitation de vitesse à 90 km/h"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4EB2
            Key             =   ""
            Object.Tag             =   "Fin de limitation de vitesse à 90 km/h"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":51CC
            Key             =   ""
            Object.Tag             =   "Début de limitation de vitesse à 110 km/h"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":54E6
            Key             =   ""
            Object.Tag             =   "Fin de limitation de vitesse à 110 km/h"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5800
            Key             =   ""
            Object.Tag             =   "Début de limitation de vitesse à 130 km/h"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5B1A
            Key             =   ""
            Object.Tag             =   "Fin de limitation de vitesse à 130 km/h"
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5E34
            Key             =   ""
            Object.Tag             =   "Entrée de bretelle"
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":614E
            Key             =   ""
            Object.Tag             =   "Sortie de bretelle"
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6468
            Key             =   ""
            Object.Tag             =   "Passage inférieur"
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6782
            Key             =   ""
            Object.Tag             =   "Passage supérieur"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6A9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":6DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":7140
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":7492
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":77E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":7B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":7E88
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":81DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":852C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":887E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":8BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":8F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":9274
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Itinéraire"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Nouveau"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileNewByImport 
         Caption         =   "Nouveau &à partir d'une campagne de mesures..."
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Ouvrir..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Fermer"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRabouter 
         Caption         =   "&Rabouter deux parcours de deux itinéraires différents..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Enre&gistrer"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Enregistrer &sous..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "Im&porter une campagne de mesures..."
      End
      Begin VB.Menu mnuViderBoitier 
         Caption         =   "&Vider le boitier de mesures..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Imprimer..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Affichage"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "Barre d'&outils"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Barre d'&état"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Fenêtre"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&Nouvelle fenêtre"
      End
      Begin VB.Menu mnuWindowBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Réorganiser les icônes"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpSommaire 
         Caption         =   "&Sommaire"
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "A&ide sur..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "À &propos de MiTemps..."
      End
      Begin VB.Menu mnuHelpBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "&Licence"
      End
   End
   Begin VB.Menu mnuRepere 
      Caption         =   "&Repere"
      Visible         =   0   'False
      Begin VB.Menu mnuRepereCreer 
         Caption         =   "&Créer un repère"
      End
      Begin VB.Menu mnuRepereSuppr 
         Caption         =   "&Supprimer le repère"
      End
      Begin VB.Menu mnuRepereBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepereDebSec 
         Caption         =   "Définir comme &début de section"
      End
      Begin VB.Menu mnuRepereFinSec 
         Caption         =   "Définir comme &fin de section"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Public Function ChoisirFichier(unTitre As String, uneExtension As String, unInitDir As String)
    'Fonction ouvrant un sélectionneur de fichier avec l'extension passée
    'en paramètre et retournant le nom complet du fichier choisi ou une
    'chaine vide si rien de choisir ou click sur Annuler
    'Le sélectionneur de fichier s'ouvre dans le répertoire unInitDir
    
    With dlgCommonDialog
        ' Active la routine de gestion d'erreur.
        On Error GoTo ErreurChoix
        
        'Ouverture d'une fenêtre Ouvrir fichier
        
        'définir les indicateurs et attributs
        'du contrôle des dialogues communs
        .CancelError = True
        .DialogTitle = unTitre
        .InitDir = unInitDir
        .Filter = uneExtension
        .flags = cdlOFNFileMustExist Or cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
        .FileName = ""
        If unTitre = MsgOpen Or unTitre = MsgImportMesure Or unTitre = MsgChoixMesure Then
            .ShowOpen
        ElseIf unTitre = MsgSaveAs Or unTitre = MsgExportTxtAs Then
            .ShowSave
        Else
            MsgBox MsgErreurProg + MsgErreurTypeShowWinInconnu + MsgIn + "frmMain:ChoisirFichier", vbCritical
        End If
        
        If Len(.FileName) = 0 Then
            'Cas où aucun fichier choisi
            ChoisirFichier = ""
        Else
            'Affectation du fichier à ouvrir
            ChoisirFichier = .FileName
        End If
        
        ' Désactive la récupération d'erreur.
        On Error GoTo 0
        'Sortie de la procédure pour éviter le passage
        'dans la gestion d'erreur
        Exit Function
    End With
    
ErreurChoix:
    'Cas où click sur Annuler
    ChoisirFichier = ""
    Exit Function
End Function


Private Sub MDIForm_Load()
    Dim unMRUSettings As Variant, unNomFich As String
    
    'Mise à jour de l'ihm du à QLM
     Call InitQlm
     
    'Affectation du fichier d'aide
    'App.HelpFile = GetAppPath() + "MiTemps.hlp"
    'dlgCommonDialog.HelpFile = App.HelpFile
    
    'Affectation du fichier d'aide
    'modif O.FOREL du 14/01/2005 :insertion du fichier chm
    unNomFich = CorrigerNomFichier(App.Path + Help_Chm)
    App.HelpFile = unNomFich
    dlgCommonDialog.HelpFile = App.HelpFile
    
    'Index des aides pour les menus
    mnuFileNew.HelpContextID = HelpID_WinNew
    mnuFileOpen.HelpContextID = HelpID_WinOpen
    mnuFileSave.HelpContextID = HelpID_WinSave
    mnuFileSaveAs.HelpContextID = HelpID_WinSaveAs
    mnuFileClose.HelpContextID = HelpID_WinClose
    mnuFileRabouter.HelpContextID = HelpID_WinRabouter
    mnuFilePrint.HelpContextID = HelpID_WinPrint
    mnuFileNewByImport.HelpContextID = HelpID_WinNewByMesure
    mnuFileImport.HelpContextID = HelpID_WinImportMesure
    mnuViderBoitier.HelpContextID = HelpID_WinViderBoitier
    mnuFileExit.HelpContextID = HelpID_WinQuit
    mnuViewToolbar.HelpContextID = HelpID_WinBarOutil
    mnuViewStatusBar.HelpContextID = HelpID_WinBarEtat
    mnuViewOptions.HelpContextID = HelpID_WinOptions
    mnuHelpAbout.HelpContextID = HelpID_WinAPropos
    
    'Récupération de la liste des fichiers récents
    unMRUSettings = GetAllSettings(App.Title, "Recent Files")
    If IsEmpty(unMRUSettings) = False Then
        'Cas où la liste des fichiers récents (MRU Files) n'est pas vide
        'getallsettings renvoit un variant non initialisé = Empty
        'On alimente les menus mnuFileMRU
        For i = UBound(unMRUSettings, 1) To 0 Step -1
            'A l'envers car on met le nom de fichier toujours en tête
            unNomFich = unMRUSettings(i, 1)
            ActualiserListeFichiersRecents unNomFich
        Next i
    End If

    'Récupération de la position et de la taille dans la base de registre
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", Screen.Width * 0.9)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", Screen.Height * 0.8)
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", (Screen.Width - Width) / 2)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", (Screen.Height - Height) / 2)
    'Me.WindowState = GetSetting(App.Title, "Settings", "WindowState", vbMaximized)
    
    'If (Screen.Width / Screen.TwipsPerPixelX) > 800 And (Screen.Height / Screen.TwipsPerPixelY) > 600 Then
        'Si résolution supérieure à 800 x 600
        'Me.WindowState = vbNormal
        'Width = Screen.Width * 0.9
        'Height = Screen.Height * 0.8
        'Top = (Screen.Height - Height) / 2
        'Left = (Screen.Width - Width) / 2
    'End If
    
    'Mise à jour des boutons dans la toolbar permettant l'impression
    'et la sauvegarde car il n'y a pas de fenêtre fille ouverte
    '==> Impression et sauvegarde impossible
    frmMain.tbToolBar.Buttons("Print").Visible = False
    frmMain.tbToolBar.Buttons("Save").Visible = False
    
    'Mise à jour des items du menu Itinéraire permettant l'impression, la fermeture
    'et la sauvegarde (save et saveas) car il n'y a pas de fenêtre fille ouverte
    '==> Impression, fermeture et sauvegarde impossible
    mnuFileSave.Enabled = False
    mnuFileSaveAs.Enabled = False
    mnuFilePrint.Enabled = False
    mnuFileClose.Enabled = False
    
    'Mise à jour de la status barre
    frmMain.sbStatusBar.Panels(1).Text = App.Title + " version " + Format(App.Major) + "." + Format(App.Minor)
    
    'Lancement d'une nouvelle étude si pas de fichier de démarrage
    'l'argument de la ligne de commande
    If Command <> "" Then
        'Ouvrir Struct-Urb avec le paramètre de la ligne commande
        '= Nom complet du fichier sur lequel on a double-cliqué
        OuvrirEtude Command
    End If
End Sub


Private Sub LoadNewDoc(Optional unNewFromMTB As Boolean = False)
    'Création d'une nouvelle fenêtre d'un nouvel itinéraire à partir de rien
    'si unNewFromMTB est false ou à partir d'un fichier MTB lu et moyenné
    'auparavant dans la fonction frmMain.mnuFileNewByImport_click
    Static lDocumentCount As Long, unPar As Parcours
    Dim frmD As frmDocument, unNbRep As Integer, unRep As Repere
    Dim unTabAbsRep As Variant, unTabTempsRep As Variant, uneDuree As Long
    Dim uneDistMax As Single, uneDureeMax As Single, uneVitMax As Single

    'Affichage du sablier en pointeur souris pour symboliser l'attente
    Me.MousePointer = vbHourglass

    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = MsgIti0 + Format(lDocumentCount)
    
    'Remplissage des libellés par défaut des conditions météo
    unePos0 = 1
    For i = 0 To 7
        'Recherche de la virgule séparant les libellés météo
        unePos = InStr(unePos0, mesOptions.mesLibMeteo, ",")
        uneStrTmp = Mid(mesOptions.mesLibMeteo, unePos0, unePos - unePos0)
        frmD.maColMeteo.Add Format(i) + " - " + uneStrTmp
        unePos0 = unePos + 1
    Next i
    uneStrTmp = Mid(mesOptions.mesLibMeteo, unePos0)
    frmD.maColMeteo.Add Format(i) + " - " + uneStrTmp
    
    If unNewFromMTB Then
        'Cas d'une création à partir d'un fichier MTB
                        
        'Remise de la souris en sablier
        MousePointer = vbHourglass
        
        'Ajout dans les parcours du nouvel itinéraire des parcours
        'issus du fichier MTB et qui ont été utilisés pour les moyennes
        For i = 1 To maColParcoursMTB.Count
            Set unPar = maColParcoursMTB(i)
            If unPar.monIsUtil Then
                frmD.maColParcours.Ajouter unPar
                'Affectation d'une couleur par défaut, on commence à 9
                'pour éviter le gris (cf aide sur fonction QBColor)
                unPar.maCouleur = QBColor(9 + frmD.maColParcours.Count Mod 6)
                'Calcul des vitesses min, max et moyenne et de la durée, du nombre
                'et du temps d'arrêts sur le parcours total
                unPar.CalculerLesVitDistDureeEtArrets -1000, 10000000
                If uneVitMax < unPar.maVmax Then uneVitMax = unPar.maVmax
                uneDuree = unPar.monTFinSection - unPar.monTDebSection
                If uneDureeMax < uneDuree Then uneDureeMax = uneDuree
                'Calcul de la distance parcourue maxi
                If uneDistMax < (unPar.maDistPar * unPar.monCoefEta) Then uneDistMax = unPar.maDistPar * unPar.monCoefEta
                'Calcul du nombre et de la durée des double tops
                'entre deux abs curv englobant tout le parcours
                unPar.CalculerNbEtDureeDoubleTop -1000, 10000000
            End If
        Next i
        
        'Calcul des durée et distance parcourue maxi
        'DonnerMaxDistDureeVit frmD.maColParcours, uneDistMax, uneDureeMax, uneVitMax
        
        'Initialisation des min et max en Distance,
        'pour avoir un bon facteur de zoom au départ pour dessiner
        'verticalement les icones des repères dans l'onglet Itinéraire
        frmD.monMinDtot = 0
        frmD.monMaxDtot = uneDistMax
        frmD.monMaxD = frmD.monMaxDtot
        frmD.monMinD = frmD.monMinDtot
        
        'Initialisation des min et max en Durée et en vitesse, pour avoir un bon
        'facteur de zoom dans les onglets visualisant la courbe Temps/Distance
        'et la courbe vitesse/temps
        frmD.monMinT = 0
        frmD.monMaxT = uneDureeMax
        frmD.monMinV = 0
        frmD.monMaxV = uneVitMax
                
        'On remet à zéro le tableau de repères
        frmD.SpreadRepere.MaxRows = 0
                        
        'Création du parcours moyen et
        'Ajout en tête dans les parcours du nouvel itinéraire
        'Le parcours moyen sera toujours celui en première position
        ' qbcolor(0) = noir
        Set unPar = New Parcours
        frmD.maColParcours.Ajouter unPar, True
        'Indication que ce parcours créé est le parcours moyen
        unPar.monIsParcoursMoyen = True
        'Récup du nombre de repère à créer = nb de valeurs moyennes
        unNbRep = maColValMoy.Count
        
        'Initialisation de l'abs curv du repère topé précédent
        uneAbsTopPred = -10000
        unNumRep = 1
        unNbDblTop = 0
        'Allocation dynamique des tableaux liés aux repères topés
        unTabAbsRep = unPar.monTabAbsRep
        unTabTempsRep = unPar.monTabTempsRep
        ReDim unTabAbsRep(1 To unNbRep)
        ReDim unTabTempsRep(1 To unNbRep)
        For i = 1 To unNbRep
            unTabAbsRep(i) = CLng(maColValMoy(i) * 10)
            '*10 Pour avoir des décimètres comme les autres parcours
            'Création des repères avec leur abs curv moyen
            'si il n'est pas trop proche du repère topé précédent
            'abs curv des repères en mètre et ecartmax en mètre
            If unTabAbsRep(i) - uneAbsTopPred > mesOptions.monEcartMax Then
                Set unRep = CreerRepere(frmD, "Repère " + Format(unNumRep), "Repère " + Format(unNumRep), CLng(unTabAbsRep(i) / 10), Autre)
                'Incrémentation du nombre de repères créés
                unNumRep = unNumRep + 1
            Else
                'Cas d'un double top, on les compte
                unNbDblTop = unNbDblTop + 1
                'Mise de l'icône jaune 'Top x2'
                'Mettre à jour l'info bulle et l'icône de l'icône repère
                ModifierIconeRepere frmD, unRep, DoubleTop
            End If
            'Stockage du repère topé précédent
            uneAbsTopPred = unTabAbsRep(i)
        Next i
        If unNbDblTop > 0 Then
            'Signalisation du nombre de double tops trouvés
            unMsg = Format(unNbDblTop) + " double tops ont été trouvés, pour chacun d'eux un seul repère, avec une icône jaune 'Top x2', sera créé dans le nouvel itinéraire."
            unMsg = unMsg + Chr(13) + Chr(13) + "Les double tops sont les repères topés séparés par moins de " + Format(mesOptions.monEcartMax) + " mètre(s), cet écart est paramètrable dans le menu Affichage/Options."
            MsgBox unMsg, vbInformation
        End If
        
        'Affectation des pointeurs sur les tableaux du parcours
        unPar.monTabAbsRep = unTabAbsRep
        unPar.monTabTempsRep = unTabTempsRep
        
        'Vidage des events pour ne voir que la fenêtre de
        'progression du calcul du parcours moyen
        DoEvents
        'Mise à jour du parcours moyen
        ActualiserParcoursMoyen unPar, frmD.maColParcours, -100, 1000000
        DoEvents
        If frmD.monMaxV < unPar.maVmax Then frmD.monMaxV = unPar.maVmax
        uneDuree = unPar.monTFinSection - unPar.monTDebSection
        If frmD.monMaxT < uneDuree Then frmD.monMaxT = uneDuree
        unPar.monIsUtil = True
        
        'Création d'un dernier repère pour la fin des mesures = distance parcourue
        'du parcours moyen convertie en mètres
        'si l'écart avec le dernier repère topé est > tolérance des options
        If (unPar.maDistPar / 10) - maColValMoy(unNbRep) > mesOptions.monEcartMax Then
            CreerRepere frmD, "Repère Fin de mesure", "Repère " + Format(unNumRep), CLng(unPar.maDistPar / 10), Autre
            unNbRep = unNumRep
        End If
        
        'Conversion du temps maxi des dixièmes de secondes en minutes
        'et de la distance maxi des décimètres en mètres
        frmD.monMaxT = frmD.monMaxT / 600
        frmD.monMaxD = frmD.monMaxD / 10
        frmD.monMaxDtot = frmD.monMaxD
        
        'Redessin en zoom total
        RedessinerZoomTout frmD
        
        'Remplissage du spread des parcours affectés
        RemplirSpreadParcours frmD
        RemplirMeteoSpreadParcours frmD
        
        'Retaillage de l'onglet itinéraire
        RetaillerOngletItiRef frmD
    Else
        'Cas d'une création à partir de rien
        
        'Récup du nombre de repère à créer
        unNbRep = 2
        
        'Mise en grisé des tous les onglets sauf le premier
        'car c'est un nouvel itinéraire que l'on ouvre
        'Les onglets vont de 0 à n-1
        For i = 1 To frmD.TabData.Tabs - 1
            frmD.TabData.TabEnabled(i) = False
        Next i
        
        'Initialisation des min et max en Distance,
        'pour avoir un bon facteur de zoom au départ
        frmD.monMinD = 0
        frmD.monMaxD = 500 ' mètres
        frmD.monMaxV = 1 ' km/h
        frmD.monMinV = 0
        frmD.monMaxT = 1 ' minute
        frmD.monMinT = 0
        
        'On remet à zéro le tableau de repères
        frmD.SpreadRepere.MaxRows = 0
        
        'Création de deux repères par défaut pour un nouvel itinéraire
        'et Dessin des icônes de repères dans la frame verticale de droite
        AfficherNouveauRepere frmD
        AfficherNouveauRepere frmD
    End If
            
    'Mise à jour des valeurs par défaut des combobox début et fin de section
    'et stockage dans le tag des dernières valeurs valides
    'Par défaut le premier et le dernier repère
    frmD.ComboRepDebSec.ListIndex = 0
    frmD.ComboRepDebSec.Tag = frmD.ComboRepDebSec.Text
    frmD.ComboRepFinSec.ListIndex = unNbRep - 1
    frmD.ComboRepFinSec.Tag = frmD.ComboRepFinSec.Text
    
    'Remplissage d'attribut par défaut
    frmD.maLongIti = frmD.maColRepere(unNbRep).monAbsCurv - frmD.maColRepere(1).monAbsCurv 'frmD.monMinD
    frmD.TextLongIti.Text = Format(frmD.maLongIti)
    frmD.monNomIti = MsgIti0 + Format(lDocumentCount)
    frmD.TextNomIti.Text = frmD.monNomIti
    
    'Affichage de la fenêtre fille
    frmD.Show
    'Affichage de l'onglet histogramme pour bien le retailler
    frmD.MSChart1.Visible = False
    'RetaillerOngletHistV frmD
    frmD.TabData.Tab = OngletHistV
    DoEvents
    'Remise en tête de l'onglet itinéraire de référence
    frmD.TabData.Tab = OngletItiRef
    frmD.MSChart1.Visible = True
    
    'Sélection graphique du repère 1, avec désélection du dernier repère créé
    DeselectionnerRepere frmD, unNbRep
    SelectionnerRepere frmD, 1
        
    'restauration du pointeur souris
    Me.MousePointer = vbDefault
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState = vbNormal Then
    'If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        SaveSetting App.Title, "Settings", "WindowState", Me.WindowState
    ElseIf Me.WindowState = vbMaximized Then
        SaveSetting App.Title, "Settings", "WindowState", Me.WindowState
    End If
    
    'Stockage en base de registre des fichiers récents
    For i = 1 To 4
        If mnuFileMRU(i - 1).Visible Then
            unFileMRU = "File" + Format(i)
            SaveSetting App.Title, "Recent Files", unFileMRU, Mid(mnuFileMRU(i - 1).Caption, 4)
        End If
    Next
End Sub





Private Sub mnuFileImport_Click()
    Dim unNomFich As String

    'Ces lignes sont commentés car si un fichier mit est ouvert
    'il est vérouillé en utilisation et ne peut être utilisé dans la fenêtre
    'd'importation
    'If Forms.Count > 1 Then
        'MsgBox "Il faut fermer toutes les fenêtres itinéraires pour pouvoir importer une campagne de mesures", vbExclamation
    'Else
    unNomFich = ChoisirFichier(MsgImportMesure, MsgMtbFile, CurDir)
    If unNomFich <> "" Then
        Me.MousePointer = vbHourglass
        'On vide les parcours issus du mtb
        ViderColParcours maColParcoursMTB
        If LireFichierMTB(unNomFich) Then
            'Cas où la lecture du MTB s'est bien passée
            'Affichage de la fenêtre de choix des parcours à importer
            frmChoixParMTB.Tag = unNomFich
            frmChoixParMTB.Show vbModal, Me
            'Affichage de la fenêtre d'importation si on clique sur le bouton
            'visualisation (= OK) et pas Annuler de frmChoixParMTB affichée
            'juste avant
            If monBtnClick = "OK" Then
                frmImportMTB.Tag = unNomFich
                frmImportMTB.Show vbModal, Me
            End If
            'On vide les parcours issus du mtb
            ViderColParcours maColParcoursMTB
        End If
        Me.MousePointer = vbDefault
    End If
    'End If
End Sub

Private Sub mnuFileNewByImport_Click()
    Dim unNomFich As String

    unNomFich = ChoisirFichier(MsgChoixMesure, MsgMtbFile, CurDir)
    If unNomFich <> "" Then
        Me.MousePointer = vbHourglass
        'On vide les parcours issus du mtb
        ViderColParcours maColParcoursMTB
        'Récup du séparateur décimale en cours
        TrouverCaractèreDécimalUtilisé
        If LireFichierMTB(unNomFich) Then
            'Cas où la lecture du MTB s'est bien passée
            'ouverture de la fenêtre de choix des parcours pour faire
            'un itinéraire de référence en moyennant les distances totales
            'et les abscisses curvilignes des repères
            frmNewParMTB.Show vbModal, Me
            'Test si la liste des parcours issus du MTB n'est pas vide
            '==> Création du nouvel itinéraire avec les moyennes
            If maColParcoursMTB.Count > 0 Then
                'Création d'un nouvel itinéraire
                LoadNewDoc True
                'On vide les parcours issus du mtb
                ViderColParcours maColParcoursMTB
            End If
        End If
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub mnuFileRabouter_Click()
    If Forms.Count > 1 Then
        MsgBox "Il faut fermer toutes les fenêtres itinéraires pour pouvoir rabouter deux parcours.", vbExclamation
    Else
        frmRabouter.Show vbModal, Me
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuLicence_Click()
    frmKey.Show 1
    'Mise à jour de l'ihm
    Call InitQlm
End Sub

Private Sub mnuRepereCreer_Click()
    AfficherNouveauRepere monIti
End Sub

Private Sub mnuRepereDebSec_Click()
    'Positionnment du début de section de travail par click droit
    Dim unNumRow As Integer
    'Récup du numéro de ligne dans le spread repère de la fenêtre fille active
    'grâce à la clé d'identification du repère dont l'icône est sélectionné
    'Cette clé a été auparavant stocké dans le tag de la MDI mère
    unNumRow = DonnerLigneRepere
    If monIti.ComboRepDebSec.ListIndex <> unNumRow - 1 Then
        'Cas où début de section change
        monIti.ComboRepDebSec.ListIndex = unNumRow - 1
        If monIti.TabData.Tab = OngletCbeDT Then
            DessinerCourbe monIti, monIti.PicBoxDT, OngletCbeDT
        ElseIf monIti.TabData.Tab = OngletCbeDV Then
            DessinerCourbe monIti, monIti.PicBoxDV, OngletCbeDV
        End If
    End If
End Sub

Private Sub mnuRepereFinSec_Click()
    'Positionnment de la fin de section de travail par click droit
    Dim unNumRow As Integer
    'Récup du numéro de ligne dans le spread repère de la fenêtre fille active
    'grâce à la clé d'identification du repère dont l'icône est sélectionné
    'Cette clé a été auparavant stocké dans le tag de la MDI mère
    unNumRow = DonnerLigneRepere
    If monIti.ComboRepFinSec.ListIndex <> unNumRow - 1 Then
        'Cas où fin de section change
        monIti.ComboRepFinSec.ListIndex = unNumRow - 1
        If monIti.TabData.Tab = OngletCbeDT Then
            DessinerCourbe monIti, monIti.PicBoxDT, OngletCbeDT
        ElseIf monIti.TabData.Tab = OngletCbeDV Then
            DessinerCourbe monIti, monIti.PicBoxDV, OngletCbeDV
        End If
    End If
End Sub

Private Sub mnuRepereSuppr_Click()
    'Suppression d'un repère à partir du menu contextuel
    's'affichant sur les icônes de repères
    Dim unNumRow As Integer
    
    'Récup du numéro de ligne dans le spread repère de la fenêtre fille active
    'grâce à la clé d'identification du repère dont l'icône est sélectionné
    'Cette clé a été auparavant stocké dans le tag de la MDI mère
    unNumRow = DonnerLigneRepere
    SupprimerRepere monIti, unNumRow
End Sub



Private Sub mnuViderBoitier_Click()
    frmViderBoitier.Show vbModal, Me
End Sub

Private Sub mnuViewOptions_Click()
    If Forms.Count > 1 Then
        MsgBox "Ouverture des options générales uniquement en consultation. Il faut fermer toutes les fenêtres itinéraires pour pouvoir changer ces options générales.", vbExclamation
    End If
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewStatusBar_Click()
    If mnuViewStatusBar.Checked Then
        sbStatusBar.Visible = False
        mnuViewStatusBar.Checked = False
    Else
        sbStatusBar.Visible = True
        mnuViewStatusBar.Checked = True
    End If
End Sub


Private Sub mnuViewToolbar_Click()
    If mnuViewToolbar.Checked Then
        tbToolBar.Visible = False
        mnuViewToolbar.Checked = False
    Else
        tbToolBar.Visible = True
        mnuViewToolbar.Checked = True
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Dim unPar As Parcours
    
    If Forms.Count > 1 Then
        'S'il y a une fenetre fille
        If monIti.maColParcours.Count > 0 Then
            'S'il y a des parcours affectés
            'Récup du parcours
            Set unPar = monIti.maColParcours(monIti.SpreadParcours.ActiveRow)
            monIti.SpreadParcours.Row = monIti.SpreadParcours.ActiveRow
            monIti.SpreadParcours.Col = monIti.SpreadParcours.ActiveCol
            
            If monIti.SpreadParcours.ActiveCol = 7 Then
                'Mise en cohérence entre la date et le jour de la semaine
                'Stockage de la date de mesure
                unPar.maDate = monIti.SpreadParcours.Text
                unJour = DonnerJourSemaine(CDate(monIti.SpreadParcours.Text))
                monIti.SpreadParcours.Col = 8
                'Mise à jour du jour de semaine
                monIti.SpreadParcours.Text = unJour
                'Stockage du jour de mesure
                unPar.monJourSemaine = monIti.SpreadParcours.Text
                'Indication du redessin de l'onglet Tableau Brut
                monIti.SetTabRedOnglet OngletTabBr, True
            End If
        End If
    End If
    
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            'On donne le focus à la MDI mère pour sortir de la saisie
            'éventuelle d'un tableau de données et on vide les événements restants
            Me.SetFocus
            DoEvents
            'Sauvegarde du fichier itinéraire en cours
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
    End Select
End Sub

'Ajout O.Forel 12/07/2005 : modification du menu aide (méthode décrites dans Chelp.bas)
Private Sub mnuHelpIndex_Click()
    Dim objHelp As CHelp
    Set objHelp = New CHelp
    'Modif fait par Frank Trifiletti on utilise le contextid de la fenêtre étude en cours
    'qui est dans la globale monetude dont son helpcontextid est mis à jour dans la sub ChangerHelpId
    'qui est appellé à chaque Form_Activate et dans le TabData_Click de frmDocument.frm
    'car le contextid était toujours nulle avec showindex normal on ne le passe pas en argument.
    If monIti Is Nothing Then
        'Cas d'appel  de F1 si aucun étude ouverte sinon plantage
        Call objHelp.ShowIndex(App.HelpFile, Help_Main)
    Else
        Call objHelp.Show(App.HelpFile, Help_Main, monIti.HelpContextID)
    End If
    'Fin modif F.Trifiletti
    Set objHelp = Nothing
End Sub

Private Sub mnuHelpSearch_Click()
    Dim objHelp As CHelp
    Set objHelp = New CHelp
    Call objHelp.ShowSearch(App.HelpFile, Help_Main)
    Set objHelp = Nothing
End Sub

Private Sub mnuHelpSommaire_Click()
    Dim objHelp As CHelp
    Set objHelp = New CHelp
    Call objHelp.Show(App.HelpFile, Help_Main)
    Set objHelp = Nothing
End Sub
'fin ajout o.Forel

'Private Sub mnuHelpContents_Click()
'    Dim nRet As Integer

    's'il n'y pas de fichier d'aide pour le projet, afficher un message à l'utilisateur
    'vous pouvez définir le fichier d'aide de votre application dans la boîte
    'de dialogue de propriétés du projet
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Impossible d'affichiez le sommaire de l'aide. Il n'y a pas d'aide associée à ce projet.", vbInformation, Me.Caption
'    Else
'        On Error Resume Next
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, &HB, 0)
'        If Err Then
'            MsgBox Err.Description
'        End If
'    End If
'End Sub

'ancien
'Private Sub mnuHelpSearch_Click()
    's'il n'y pas de fichier d'aide pour le projet, afficher un message à l'utilisateur
    'vous pouvez définir le fichier d'aide de votre application dans la boîte
    'de dialogue de propriétés du projet
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Impossible d'afficher le sommaire de l'aide. Il n'y a pas d'aide associée à ce projet.", vbInformation, Me.Caption
'    Else
'        If HelpContextID = 0 Then
'            'Lance le sommaire de l'aide
'            mnuHelpContents_Click
'        Else
'            'Lance l'aide du bon contexte
'            dlgCommonDialog.HelpCommand = cdlHelpContext
'            dlgCommonDialog.HelpContext = HelpContextID
'            dlgCommonDialog.ShowHelp  ' affiche la rubrique
'        End If
'    End If
'End Sub


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub


Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub


Private Sub mnuWindowNewWindow_Click()
    mnuFileNew_Click
End Sub



Private Sub mnuFileOpen_Click()
    Dim unNomFich As String
    Dim uneForm As Form, unIsMDIForm As Boolean

    unNomFich = ChoisirFichier(MsgOpen, MsgMitFile, CurDir)
    If unNomFich <> "" Then
        Me.MousePointer = vbHourglass
        'Récupération de la fenêtre active et on la rend inactive
        'pour éviter que le 2ème click du double click sur un fichier mit
        'déclenche une sélection si les onglets de visu graphique sont au
        'premier plan
        unIsMDIForm = (Screen.ActiveForm Is frmMain)
        If unIsMDIForm = False Then
            Set uneForm = Screen.ActiveForm
            uneForm.Enabled = False
        End If
        'Ouverture de l'étude
        OuvrirEtude unNomFich
        If unIsMDIForm = False Then
            'On remet la form inactive en active
            uneForm.Enabled = True
        End If
        Me.MousePointer = vbDefault
    End If
End Sub


Private Sub mnuFileClose_Click()
    Unload monIti
End Sub


Private Sub mnuFileSave_Click()
    'Sauvegarde de l'étude active
    'Le nom du fichier ne sert que si c'est un itinéraire existant
    'Titre fenetre = "Itinéraire " + numéro ou nom fichier
    SauverFichier monIti, Mid(monIti.Caption, 12), False
End Sub


Private Sub mnuFileSaveAs_Click()
    'Sauvegarde de l'étude active
    'Le nom du fichier est vide car on fait un enregistrer sous
    'le nom de fichier est choisi dans SauverFichier
    SauverFichier monIti, "", True
End Sub


Private Sub mnuFilePrint_Click()
    'Si protection invalide on ne fait rien
    'If ProtectCheck(2) <> 0 Then
    '    Exit Sub
    If Printers.Count = 0 Then
        MsgBox "Aucune imprimante n'est connectée à ce poste", vbCritical
    ElseIf monIti.maColParcours.Count = 0 Then
        MsgBox "Impression impossible : L'itinéraire ouvert (" + monIti.Caption + ") ne contient aucun parcours.", vbExclamation
    ElseIf DonnerNbParcoursUtil(monIti) = 0 Then
        MsgBox "Impression impossible : L'itinéraire ouvert (" + monIti.Caption + ") n' a aucun parcours utilisé.", vbExclamation
    Else
        frmImprimer.Show vbModal
    End If
End Sub


Private Sub mnuFileMRU_Click(Index As Integer)
    OuvrirEtude Mid(mnuFileMRU(Index).Caption, 4)
End Sub


Private Sub mnuFileExit_Click()
    'décharger la feuille
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub


Public Sub AfficherMenuContextuel(unIndRepere As Integer)
    'Affichage du menu contextuel des repères de la frame verticale
    'contenant les icones de repères de la fenêtre active
    
    'Récupération du nom court du repère
    unNomRep = monIti.maColRepere(monIti.ImageRepere(unIndRepere).Tag).monNomCourt
    
    'Affichage d'un menu contextuel en modifiant ces items (chr(34) = ")
    frmMain.mnuRepereCreer.Enabled = (monIti.CheckSection = 0)
    frmMain.mnuRepereSuppr.Enabled = (monIti.CheckSection = 0)
    frmMain.mnuRepereSuppr.Caption = "&Supprimer le repère " + Chr(34) + unNomRep + Chr(34)
    frmMain.mnuRepereDebSec.Enabled = (monIti.CheckSection = 1)
    frmMain.mnuRepereDebSec.Caption = "Définir " + Chr(34) + unNomRep + Chr(34) + " comme &début de section"
    frmMain.mnuRepereFinSec.Enabled = (monIti.CheckSection = 1)
    frmMain.mnuRepereFinSec.Caption = "Définir " + Chr(34) + unNomRep + Chr(34) + " comme &fin de section"
    PopupMenu frmMain.mnuRepere, vbPopupMenuRightButton
End Sub

'Code pour modifier l'ihm suite à l'implémentation de Qlm
Private Sub InitQlm()
    'Initialisation des menus modifiés par QLM
    'les variables globales sont maj par protection.bas
    'ATTENTION : vérifier les noms des menus!!!
    Me.mnuHelpBar2.Visible = GvisibiliteMnuBarre
    Me.mnuLicence.Visible = GvisibiliteMnuLicence
    'a adapter en fonction du clogiciel
    Me.Caption = "MiTemps v" + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision) + GmodifTitreApplication
    'fin initialisation qlm
    'fin initialisation qlm
End Sub
