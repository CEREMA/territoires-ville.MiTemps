VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Begin VB.Form frmImprimer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimer"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   Icon            =   "frmImprimer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleMode       =   0  'User
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameEchImp 
      Height          =   2175
      Left            =   60
      TabIndex        =   18
      Top             =   4440
      Width           =   6855
      Begin FPSpread.vaSpread SpreadEchImp 
         Height          =   1455
         Left            =   120
         TabIndex        =   20
         Top             =   660
         Width           =   6615
         _Version        =   131077
         _ExtentX        =   11668
         _ExtentY        =   2566
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
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
         MaxCols         =   3
         MaxRows         =   4
         ProcessTab      =   -1  'True
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmImprimer.frx":0442
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin VB.CheckBox CheckModifEchImp 
         Caption         =   "Modifier les échelles d'impressions automatique pour les courbes Distance/Temps, Distance/Vitesse et le synoptique des vitesses"
         Height          =   400
         Left            =   120
         TabIndex        =   19
         Top             =   180
         Value           =   2  'Grayed
         Width           =   6615
      End
   End
   Begin VB.Frame FrameEltPrint 
      Caption         =   "Eléments à imprimer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   60
      TabIndex        =   8
      Top             =   2400
      Width           =   5415
      Begin VB.CheckBox CheckTabSynStat 
         Caption         =   "Tableau Synthèse / Statisitiques"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CheckBox CheckTabBrut 
         Caption         =   "Tableau Brut"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox CheckHistoV 
         Caption         =   "Histogramme des vitesses"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox CheckSynoV 
         Caption         =   "Synoptique des vitesses"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox CheckCourbeDV 
         Caption         =   "Courbe Distance/Vitesse"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox CheckCourbeDT 
         Caption         =   "Courbe Distance/Temps"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox CheckNomIti 
         Caption         =   "Nom de l'itinéraire"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox CheckNomFich 
         Caption         =   "Nom du fichier Itinéraire"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin VB.Frame FrameInfoImp 
      Height          =   2295
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   5415
      Begin VB.ComboBox ComboEpTrait 
         Height          =   315
         ItemData        =   "frmImprimer.frx":0A21
         Left            =   1800
         List            =   "frmImprimer.frx":0A34
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label LabelOrient 
         AutoSize        =   -1  'True
         Caption         =   "Type Orientation"
         Height          =   195
         Left            =   2160
         TabIndex        =   17
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Epaisseur de trait :"
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
         TabIndex        =   7
         Top             =   1680
         Width           =   1620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Orientation :"
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
         TabIndex        =   6
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label NomImp 
         AutoSize        =   -1  'True
         Caption         =   "Imprimante courante : "
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
         TabIndex        =   5
         Top             =   240
         Width           =   1920
      End
      Begin VB.Image ImagePaysage 
         Height          =   495
         Left            =   1320
         Picture         =   "frmImprimer.frx":0A47
         Stretch         =   -1  'True
         Top             =   720
         Width           =   600
      End
      Begin VB.Image ImagePortrait 
         Height          =   600
         Left            =   1320
         Picture         =   "frmImprimer.frx":1189
         Stretch         =   -1  'True
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5580
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5580
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "Configurer imprimante..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5580
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "frmImprimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable stockant les modifications d'échelle pour l'impression
'des courbes Distance/Temps, Distance/Vitesse, Synoptique des vitesses
Private maModifEchImp As Integer
Private monMinDImp As Single
Private monMaxDImp As Single
Private monMinTImp As Single
Private monMaxTImp As Single
Private monMinVImp As Single
Private monMaxVImp As Single

Private Sub CheckModifEchImp_Click()
    'Affichage ou masquage du tableau permettant d'être en modification
    'd'échelle pour les impressions.
    Dim unNumCol As Integer
    
    'A l'ouverture d'une fenêtre itinéraire monIti.maModifEchImp =0, puis
    'à chaque cochage ou décochage monIti.maModifEchImp est incrémentée de 1,
    'donc monIti.maModifEchImp est paire si la checkbox de modif d'echelle en
    'impression est cochée et impaire si elle est décochée
    maModifEchImp = maModifEchImp + 1
    
    If CheckModifEchImp.Value = 0 Then
        'Case de modif des échelles en impression décochée
        'Masquage du tableau des modifications des échelles en impression
        SpreadEchImp.Visible = False
        FrameEchImp.Height = 660 'juste pour voir la case à cocher
    Else
        'Case de modif des échelles en impression cochée
        'Affichage du tableau des modifications des échelles en impression
        SpreadEchImp.Visible = True
        FrameEchImp.Height = SpreadEchImp.Top + SpreadEchImp.Height + 120
        'Affichage du tableau des min/max en échelle avec le séparateur
        'décimal en cours d'utilisation
        TrouverCaractèreDécimalUtilisé
        For unNumCol = 1 To SpreadEchImp.MaxCols
            InitColSpreadCaractèreDécimal SpreadEchImp, unNumCol, monCarDeci
        Next unNumCol
        RemplirTabModifEchImp
End If
    'Retaillage de la fenêtre d'impression
    uneLargeurBandeauTitreFenetre = frmImprimer.Height - frmImprimer.ScaleHeight
    frmImprimer.Height = uneLargeurBandeauTitreFenetre + FrameEchImp.Top + FrameEchImp.Height + 120
End Sub


Private Sub cmdCancel_Click()
    FermerFenetre Me
End Sub

Private Sub cmdOK_Click()
    Dim uneMargeG As Single, uneMargeD As Single
    Dim uneMargeH As Single, uneMargeB As Single
    Dim uneLargeur As Single, uneHauteur As Single
    Dim unNumPage As Byte, unHeader As String
    Dim uneHText As Integer, uneWText As Integer
    Dim unMinDSave As Single, unMaxDSave As Single
    Dim unMinTSave As Single, unMaxTSave As Single
    Dim unMinVSave As Single, unMaxVSave As Single
    
    'Affichage du sablier en pointeur souris pour symboliser l'attente
    Me.MousePointer = vbHourglass
    unNumPage = 0
    uneHText = Printer.TextHeight("H")
    uneWText = Printer.TextWidth("W")
    
    'Stockage tant que la fenêtre itinéraire active est ouverte du choix
    'de modifications d'échelle en impression
    monIti.maModifEchImp = maModifEchImp
    monIti.monMinDImp = monMinDImp
    monIti.monMaxDImp = monMaxDImp
    monIti.monMinTImp = monMinTImp
    monIti.monMaxTImp = monMaxTImp
    monIti.monMinVImp = monMinVImp
    monIti.monMaxVImp = monMaxVImp
    
    'Modification des échelles pour les impressions si voulu
    'par l'utilisateur
    If maModifEchImp Mod 2 = 1 Then
        'Si impaire, la case à cocher "modif échelle impression" est cochée
        'Vérification que les min user <= au min auto
        'et que les max user >= max auto
        If monIti.monMinDImp > monIti.monMinD + 0.01 Or monIti.monMinTImp > monIti.monMinT + 0.01 Or monIti.monMinVImp > monIti.monMinV + 0.01 Then
            MsgBox "Les minimums modifiables doivent être inférieurs ou égaux aux minimums automatiques.", vbInformation
            Me.MousePointer = vbDefault
            Exit Sub
        ElseIf monIti.monMaxDImp < monIti.monMaxD - 0.01 Or monIti.monMaxTImp < monIti.monMaxT - 0.01 Or monIti.monMaxVImp < monIti.monMaxV - 0.01 Then
            MsgBox "Les maximums modifiables doivent être supérieurs ou égaux aux maximums automatiques.", vbInformation
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        'Sauvegarde des min/max en distance, temps et vitesse pour les dessins
        'dans les onglets distance/temps, distance/vitesse et le synoptique
        'des vitesses
        unMinDSave = monIti.monMinD
        unMaxDSave = monIti.monMaxD
        unMinTSave = monIti.monMinT
        unMaxTSave = monIti.monMaxT
        unMinVSave = monIti.monMinV
        unMaxVSave = monIti.monMaxV
        'Modification des min/max en distance, temps et vitesse pour les courbes
        'distance/temps, distance/vitesse et le synoptique des vitesses
        monIti.monMinD = monMinDImp
        monIti.monMaxD = monMaxDImp
        monIti.monMinT = monMinTImp
        monIti.monMaxT = monMaxTImp
        monIti.monMinV = monMinVImp
        monIti.monMaxV = monMaxVImp
    End If
    
    'Stockage dans la base de registre de l'épaisseur de trait
    mesOptions.monEpaisTrait = Val(ComboEpTrait.Text)
    SaveSetting App.Title, "Options", "EpaisTrait", mesOptions.monEpaisTrait
    
    'Initialisation du printer
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 7
    
    'Remise en trait noir pas gras
    Printer.Font.Bold = False
    Printer.ForeColor = QBColor(0)
    
    'Dessin de la courbe DT
    If CheckCourbeDT.Value = 1 Then
        ImprimerEnteteEtLegende unNumPage, "Courbe Distance/Temps"
        DessinerCourbe monIti, Printer, OngletCbeDT
    End If
    
    'Dessin de la courbe DV
    If CheckCourbeDV.Value = 1 Then
        ImprimerEnteteEtLegende unNumPage, "Courbe Distance/Vitesse"
        DessinerCourbe monIti, Printer, OngletCbeDV
    End If
    
    'Dessin du synoptique des vitesses
    If CheckSynoV.Value = 1 Then
        ImprimerEnteteEtLegende unNumPage, "Synoptique des vitesses"
        'Affichage de la légende juste après la marge du haut
        'de la fenêtre itinéraire
        unCurY = Printer.CurrentY
        uneMargeG = 0.9 * UnCmEnTwips
        Printer.CurrentX = uneMargeG
        Printer.CurrentY = unCurY
        Printer.Font.Bold = True
        Printer.Print "Légende (V en km/h): "
        uneString = "Légende (V en km/h): "
        
        Printer.ForeColor = mesOptions.maCouleurClasV1
        Printer.CurrentX = uneMargeG + TextWidth(uneString)
        Printer.CurrentY = unCurY
        uneStrTmp = "0 <= V <= " + Format(mesOptions.maValClasV1) + "  "
        Printer.Print uneStrTmp
        Printer.Line (uneMargeG + TextWidth(uneString), unCurY + uneHText)-(uneMargeG + TextWidth(uneString + uneStrTmp) - uneWText, unCurY + uneHText * 2), Printer.ForeColor, BF
        
        uneString = uneString + uneStrTmp
        Printer.ForeColor = mesOptions.maCouleurClasV2
        Printer.CurrentX = uneMargeG + TextWidth(uneString)
        Printer.CurrentY = unCurY
        uneStrTmp = Format(mesOptions.maValClasV1) + " < V <= " + Format(mesOptions.maValClasV2) + "  "
        Printer.Print uneStrTmp
        Printer.Line (uneMargeG + TextWidth(uneString), unCurY + uneHText)-(uneMargeG + TextWidth(uneString + uneStrTmp) - uneWText, unCurY + uneHText * 2), Printer.ForeColor, BF
        
        uneString = uneString + uneStrTmp
        Printer.ForeColor = mesOptions.maCouleurClasV3
        Printer.CurrentX = uneMargeG + TextWidth(uneString)
        Printer.CurrentY = unCurY
        uneStrTmp = Format(mesOptions.maValClasV2) + " < V <= " + Format(mesOptions.maValClasV3) + "  "
        Printer.Print uneStrTmp
        Printer.Line (uneMargeG + TextWidth(uneString), unCurY + uneHText)-(uneMargeG + TextWidth(uneString + uneStrTmp) - uneWText, unCurY + uneHText * 2), Printer.ForeColor, BF
        
        uneString = uneString + uneStrTmp
        Printer.ForeColor = mesOptions.maCouleurClasV4
        Printer.CurrentX = uneMargeG + TextWidth(uneString)
        Printer.CurrentY = unCurY
        uneStrTmp = Format(mesOptions.maValClasV3) + " < V <= " + Format(mesOptions.maValClasV4) + "  "
        Printer.Print uneStrTmp
        Printer.Line (uneMargeG + TextWidth(uneString), unCurY + uneHText)-(uneMargeG + TextWidth(uneString + uneStrTmp) - uneWText, unCurY + uneHText * 2), Printer.ForeColor, BF
        
        uneString = uneString + uneStrTmp
        Printer.ForeColor = mesOptions.maCouleurClasV5
        Printer.CurrentX = uneMargeG + TextWidth(uneString)
        Printer.CurrentY = unCurY
        uneStrTmp = Format(mesOptions.maValClasV4) + " < V <= " + Format(mesOptions.maValClasV5) + "  "
        Printer.Print uneStrTmp
        Printer.Line (uneMargeG + TextWidth(uneString), unCurY + uneHText)-(uneMargeG + TextWidth(uneString + uneStrTmp) - uneWText, unCurY + uneHText * 2), Printer.ForeColor, BF
        
        uneString = uneString + uneStrTmp
        Printer.ForeColor = mesOptions.maCouleurClasV6
        Printer.CurrentX = uneMargeG + TextWidth(uneString)
        Printer.CurrentY = unCurY
        uneStrTmp = "V > " + Format(mesOptions.maValClasV5)
        Printer.Print uneStrTmp
        Printer.Line (uneMargeG + TextWidth(uneString), unCurY + uneHText)-(uneMargeG + TextWidth(uneString + uneStrTmp) - uneWText, unCurY + uneHText * 2), Printer.ForeColor, BF
        
        'Stockage de la marge du haut de la fenêtre itinéraire avant modif
        uneMargeHautSave = monIti.maMargeHaut
        'Modif de la marge du haut de la fenêtre itinéraire
        monIti.maMargeHaut = unCurY + Printer.TextHeight("W") * 4
        'Dessin du synoptique des vitesses
        Printer.ForeColor = QBColor(0)
        Printer.Font.Bold = False
        DessinerSynoV monIti, Printer
        'Restauration de la marge du haut de la fenêtre itinéraire
        monIti.maMargeHaut = uneMargeHautSave
    End If
    
    'Dessin de l'histogramme des vitesses
    If CheckHistoV.Value = 1 Then
        ImprimerEnteteEtLegende unNumPage, "Histogramme des vitesses", True
        'Redessin du synotique des vitesses
        DessinerHistoV monIti
        DoEvents
        'On vide d'abord le Presse-Papier
        Clipboard.Clear
        'Masquage de la légende
        monIti.MSChart1.ShowLegend = False
        'Copie de l'image de l'histogramme dans le Presse-Papier au format wmf
        monIti.MSChart1.EditCopy
        'Récup de l'image de format wmf dans le Presse-Papier
        Set unePicture = Clipboard.GetData(vbCFMetafile)
        'Impression de l'image
        uneMargeG = 0.5 * UnCmEnTwips
        uneMargeB = 2 * UnCmEnTwips
        uneMargeD = 2 * UnCmEnTwips
        uneMargeH = Printer.CurrentY
        uneLargeur = Printer.Width - uneMargeG - uneMargeD
        uneHauteur = Printer.Height - uneMargeH - uneMargeB
        'Recherche du coté de taille mini pour avoir un agrandissement carré
        If uneHauteur > uneLargeur Then
            unCoteMin = uneLargeur
        Else
            unCoteMin = uneHauteur
        End If
        Printer.PaintPicture unePicture, uneMargeG, uneMargeH, unCoteMin, unCoteMin
        'Affichage de la légende
        monIti.MSChart1.ShowLegend = True
    End If
    
    'Dessin du tableau brut
    If CheckTabBrut.Value = 1 Then
        RemplirTabBrut monIti
        unHeader = "/fb1" + App.Title + " version " + Format(App.Major) + "." + Format(App.Minor)
        'Impression du nom de fichier et du nom de l'itinéraire
        If CheckNomFich.Value = 1 Then
            unHeader = unHeader + "/n" + "Fichier : " + monIti.Caption
        End If
        If CheckNomIti.Value = 1 Then
            unHeader = unHeader + "/n" + "Itinéraire : " + monIti.monNomIti
        End If
        unHeader = unHeader + "/n/n" + "Tableau brut"
        ConfigurerSpreadToPrint monIti.SpreadTabBrut, unHeader, "Tableau brut"
        'Sélection de la ligne d'entête pour mettre une fonte grasse
        monIti.SpreadTabBrut.Row = 0
        monIti.SpreadTabBrut.Col = -1
        monIti.SpreadTabBrut.FontBold = True
        monIti.SpreadTabBrut.Action = 13 'SS_ACTION_PRINT
        'Remise de la ligne d'entête en fonte simple
        monIti.SpreadTabBrut.Row = 0
        monIti.SpreadTabBrut.Col = -1
        monIti.SpreadTabBrut.FontBold = False
    End If
    
    'Dessin du tableau de synthèse et de statistiques
    If CheckTabSynStat.Value = 1 Then
        RemplirTabSS monIti
        unHeader = "/fb1" + App.Title + " version " + Format(App.Major) + "." + Format(App.Minor)
        'Impression du nom de fichier et du nom de l'itinéraire
        If CheckNomFich.Value = 1 Then
            unHeader = unHeader + "/n" + "Fichier : " + monIti.Caption
        End If
        If CheckNomIti.Value = 1 Then
            unHeader = unHeader + "/n" + "Itinéraire : " + monIti.monNomIti
        End If
        unHeader = unHeader + "/n/n" + "Tableau de synthèse et de statistiques avec " + Format(DonnerNbParcoursUtil(monIti)) + " parcours pris en compte"
        ConfigurerSpreadToPrint monIti.SpreadTabSS, unHeader, "Tableau de synthèse et de statistiques"
        'Sélection de la ligne d'entête pour mettre une fonte grasse
        monIti.SpreadTabSS.Row = 0
        monIti.SpreadTabSS.Col = -1
        monIti.SpreadTabSS.FontBold = True
        monIti.SpreadTabSS.Action = 13 'SS_ACTION_PRINT
        'Remise de la ligne d'entête en fonte simple
        monIti.SpreadTabSS.Row = 0
        monIti.SpreadTabSS.Col = -1
        monIti.SpreadTabSS.FontBold = False
    End If
    
    'Envoi à l'imprimante
    Printer.EndDoc
    
    'Restauration du pointeur souris par défaut
    DoEvents
    Me.MousePointer = vbDefault
    
    'Restauration des min/max en distance, temps et vitesse pour les dessins
    'dans les onglets distance/temps, distance/vitesse et le synoptique
    'des vitesses s'il y a eu modification des échelles pour les impressions
    'par l'utilisateur
    If maModifEchImp Mod 2 = 1 Then
        monIti.monMinD = unMinDSave
        monIti.monMaxD = unMaxDSave
        monIti.monMinT = unMinTSave
        monIti.monMaxT = unMaxTSave
        monIti.monMinV = unMinVSave
        monIti.monMaxV = unMaxVSave
    End If
    
    'Fermeture de la fenêtre d'impression
    FermerFenetre Me
End Sub


Private Sub cmdConfig_Click()
    Caption = UCase(MsgWaitConfig)
    'Affichage de la fenetre de configuration d'imprimante
    'par appel de celle de la dll, comdlg32.dll
    'car comdlg32.ocx est buggé en W95,98 (printer par défaut changé après son
    'utilisation) et sous NT (orientation, taille papier inchangeable)
    'ShowPrinter fonction défini dans ModulePrintAPI.bas
    ShowPrinter Me, PD_PRINTSETUP
    
    'If PlateformeNT Then
        'Affichage de la fenetre de configuration d'imprimante
        'par appel de celle de la dll, comdlg32.dll
        'car comdlg32.ocx est buggé en W95,98
        '(orientation, taille papier inchangeable)
        'ShowPrinter fonction défini dans ModulePrintAPI.bas
        'ShowPrinter Me, PD_PRINTSETUP
    'Else
        ' Active la routine de gestion d'erreur.
        'On Error GoTo CancelPress
        
        'Affichage de la fenetre de configuration d'imprimante
        'frmMain.dlgCommonDialog.CancelError = True
        'frmMain.dlgCommonDialog.flags = cdlPDPrintSetup
        'frmMain.dlgCommonDialog.ShowPrinter
    'End If
    
    DoEvents
    Caption = MsgTitreFrmImp
    'Mise à jour du nom de l'imprimante courante
    NomImp.Caption = "Imprimante courante : " + Printer.DeviceName
    'Mise à jour de l'orientation
    If Printer.Orientation = vbPRORPortrait Then
        'Cas d'une orientation portrait
        ImagePortrait.Visible = True
        ImagePaysage.Visible = False
        LabelOrient.Caption = "Portrait"
    Else
        'Cas d'une orientation paysage
        ImagePortrait.Visible = False
        ImagePaysage.Visible = True
        LabelOrient.Caption = "Paysage"
    End If
    
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    Exit Sub 'Pour éviter le traitement des erreurs s'il n'y a pas eu
    
    'Gestion des erreurs
CancelPress:
    
    Caption = MsgTitreFrmImp
    Select Case Err.Number
        Case cdlCancel 'Click sur le bouton Annuler
            'On ne fait rien
        Case Else
            ' Traite les autres situations ici...
            unMsg = "Erreur " + Format(Err.Number) + " : " + Err.Description
            MsgBox unMsg, vbCritical
    End Select
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not ActiveControl Is SpreadEchImp And KeyCode = vbKeyReturn Then
        cmdOK_Click
    End If
End Sub

Private Sub Form_Load()
    InitialiserPrinter
    NomImp.Caption = "Imprimante courante : " + Printer.DeviceName
    Caption = MsgTitreFrmImp
    CentrerFenetreEcran Me
    
    'Contexte d'aide
    HelpContextID = HelpID_WinPrint
    
    'Mise à jour de l'orientation
    If Printer.Orientation = vbPRORPortrait Then
        'Cas d'une orientation portrait
        ImagePortrait.Visible = True
        ImagePaysage.Visible = False
        LabelOrient.Caption = "Portrait"
    Else
        'Cas d'une orientation paysage
        ImagePortrait.Visible = False
        ImagePaysage.Visible = True
        LabelOrient.Caption = "Paysage"
    End If
    
    'Affichage de l'épaisseur de trait
    ComboEpTrait.Text = Format(mesOptions.monEpaisTrait)
    
    'Affichage évenuelle des modifications d'échelle pour les impressions
    'A l'ouverture d'une fenêtre itinéraire monIti.maModifEchImp =0, puis
    'à chaque cochage ou décochage monIti.maModifEchImp est incrémentée de 1,
    'donc monIti.maModifEchImp est paire si la checkbox de modif d'echelle en
    'impression est décochée et impaire si elle est cochée
    maModifEchImp = monIti.maModifEchImp + 1
    '+1 car on fait tjs + 1 dans le CheckModifEchImp.value, parité inchangée
    'aprés le CheckModifEchImp.Value ci-dessous
    CheckModifEchImp.Value = monIti.maModifEchImp Mod 2
    'CheckModifEchImp.value vaudra 0 ou 1 or elle est initialisée à grisée = 2
    'pour être sur de passer dans le click event de CheckModifEchImp et faire
    'le traitement associé
    If monIti.maModifEchImp = 0 Then
        'Initialisation des min/max impression en distance, vitesse et temps
        'avec les valeurs des min/max automatique, juste la première fois
        monIti.monMinDImp = monIti.monMinD
        monIti.monMaxDImp = monIti.monMaxD
        monIti.monMinTImp = monIti.monMinT
        monIti.monMaxTImp = monIti.monMaxT
        monIti.monMinVImp = monIti.monMinV
        monIti.monMaxVImp = monIti.monMaxV
    End If
    'Initialisation des min/max impression en distance, vitesse et temps de la
    'fenêtre avec les valeurs des min/max impression de la fenêtre itinéraire
    monMinDImp = monIti.monMinDImp
    monMaxDImp = monIti.monMaxDImp
    monMinTImp = monIti.monMinTImp
    monMaxTImp = monIti.monMaxTImp
    monMinVImp = monIti.monMinVImp
    monMaxVImp = monIti.monMaxVImp
End Sub

Public Sub ImprimerEnteteEtLegende(unNumPage As Byte, unTitre As String, Optional unPrintHistoV As Boolean = False)
    'Procédure imprimant l'entête et la légende des couleurs des parcours
    'Impression du nom du fichier et du nom de l'itinéraire
    'Cette procédure met à jour la largeur de l'entête = la marge du haut
    Dim unPar As Parcours, uneMargeG As Single, unNomFich As String
    Dim unCurY As Single, unText As String, unNbLignes As Byte
    
    If unNumPage > 0 Then
        Printer.NewPage
    Else
        unNumPage = 1
    End If
    
    'Initialisation des marges
    uneMargeG = 0.9 * UnCmEnTwips
    uneMargeHB = 0.5 * UnCmEnTwips
    
    'Impression du logiciel et de sa version
    Printer.Font.Bold = True
    Printer.DrawStyle = vbSolid
    Printer.ForeColor = QBColor(0) 'dessin en noir
    Printer.CurrentX = uneMargeG
    Printer.CurrentY = uneMargeHB
    Printer.Print App.Title + " version " + Format(App.Major) + "." + Format(App.Minor)
    
    'Impression du nom de fichier et du nom de l'itinéraire
    If CheckNomFich.Value = 1 Then
        Printer.CurrentX = uneMargeG
        If EstNouvelIti(monIti) Then
            unNomFich = monIti.Caption
        Else
            unNomFich = Mid(monIti.Caption, 12)
        End If
        Printer.Print "Fichier : " + unNomFich
    End If
    If CheckNomIti.Value = 1 Then
        uneFtSize = Printer.Font.Size
        'Taille de la fonte * 1.5 pour le titre de l'itinéraire
        Printer.Font.Size = uneFtSize * 1.5
        Printer.CurrentX = uneMargeG
        Printer.Print "Itinéraire : " + monIti.monNomIti
        'Remise de la fonte à la taille initiale
        Printer.Font.Size = uneFtSize
    End If
    
    'Impression du tableau des info des parcours sélectionnés
    unCurY = Printer.CurrentY + 20
    unCurY0 = Printer.CurrentY + 10
    unDecH = Printer.TextHeight("W")
    'Dessin de la ligne plafond du tableau
    Printer.Line (uneMargeG - 30, unCurY0)-(uneMargeG - 30 + UnCmEnTwips * 18.6, unCurY0), 0
    'Impression de la ligne d'entête
    ImprimerTexte uneMargeG, unCurY, "Parcours"
    If unPrintHistoV = False Then
        ImprimerTexte uneMargeG + UnCmEnTwips * 2.6, unCurY, "N°" 'Pour mettre 20 fois W en majuscule
    Else
        'Impression de l'histogramme des vitesses
        ImprimerTexte uneMargeG + UnCmEnTwips * 2.6, unCurY, "" 'Pour mettre 20 fois W en majuscule
    End If
    ImprimerTexte uneMargeG + UnCmEnTwips * 3.1, unCurY, "Date"
    ImprimerTexte uneMargeG + UnCmEnTwips * 4.47, unCurY, "Heure"
    ImprimerTexte uneMargeG + UnCmEnTwips * 5.25, unCurY, "Vmoy(km/h)"
    ImprimerTexte uneMargeG + UnCmEnTwips * 6.75, unCurY, "Vmin(km/h)"
    ImprimerTexte uneMargeG + UnCmEnTwips * 8.2, unCurY, "Vmax(km/h)"
    ImprimerTexte uneMargeG + UnCmEnTwips * 9.7, unCurY, "Nb arrêts"
    ImprimerTexte uneMargeG + UnCmEnTwips * 10.9, unCurY, "Temps arrêts"
    ImprimerTexte uneMargeG + UnCmEnTwips * 12.55, unCurY, "Nb dbl Top"
    ImprimerTexte uneMargeG + UnCmEnTwips * 13.95, unCurY, "Tps dbl Top"
    ImprimerTexte uneMargeG + UnCmEnTwips * 15.6, unCurY, "Durée totale"
    ImprimerTexte uneMargeG + UnCmEnTwips * 17.25, unCurY, "Dist tot(m)"
    'Dessin de la ligne plancher de la ligne i du tableau
    Printer.Line (uneMargeG - 30, unCurY + unDecH + 10)-(uneMargeG - 30 + UnCmEnTwips * 18.6, unCurY + unDecH + 10), 0
    'Dessin de la ligne vertical fermant la colonne
    Printer.Line (uneMargeG - 30 + UnCmEnTwips * 18.6, unCurY0)-(uneMargeG - 30 + UnCmEnTwips * 18.6, unCurY + unDecH + 10), 0
    
    unCurY = unCurY + unDecH + 20
    Printer.Font.Bold = False
    unDecH = Printer.TextHeight("W")
    For i = 1 To monIti.maColParcours.Count
        'Récup du parcours
        Set unPar = monIti.maColParcours(i)
        'Cas où le parcours est utilisé ==> Mis dans le tableau légende
        If unPar.monIsUtil Then
            If Printer.TextWidth(unPar.monNom) > UnCmEnTwips * 2.6 Then
                'Si nom trop grand on le met sur 2 lignes
                unDecH0 = unDecH
                unNbLignes = 2
                ImprimerTexte uneMargeG, unCurY, Mid(unPar.monNom, 1, 10), unNbLignes
                ImprimerTexte uneMargeG, unCurY + unDecH0 + 10, Mid(unPar.monNom, 11), unNbLignes
            Else
                unNbLignes = 1
                unDecH0 = 0
                ImprimerTexte uneMargeG, unCurY, unPar.monNom, unNbLignes
            End If
            'Dessin de la ligne en niveau de gris pour les impressions
            'noir et blanc sous le nom du parcours
            'Test si l'imprimante est noir et blanc ou couleur
            'unPrinterNB = Vrai ===> N&B, unPrinterNB = Faux ===> Couleur
            unPrinterNB = (Printer.ColorMode <> vbPRCMColor)
            If unPrinterNB Then
                Printer.DrawStyle = vbSolid
                Printer.Line (uneMargeG - 30, unCurY + unDecH + unDecH0 - 20)-(uneMargeG - 30 + UnCmEnTwips * 2.6, unCurY + unDecH + unDecH0 - 20), unPar.maCouleur
            End If
            
            'Signalétique de la couleur du parcours
            Printer.ForeColor = unPar.maCouleur
            If unPrintHistoV = False Then
                ImprimerTexte uneMargeG + UnCmEnTwips * 2.6, unCurY, "P" + Format(i - 1), unNbLignes 'Pour mettre 20 fois W en majuscule sur deux lignes
            Else
                'Impression de l'histogramme des vitesses
                ImprimerTexte uneMargeG + UnCmEnTwips * 2.6, unCurY, "", unNbLignes 'Pour mettre 20 fois W en majuscule sur deux lignes
                Printer.Line (uneMargeG + UnCmEnTwips * 2.6, unCurY)-(uneMargeG + UnCmEnTwips * 3, unCurY + unDecH * unNbLignes), Printer.ForeColor, BF
            End If
            Printer.ForeColor = QBColor(0) 'Remise en noir
            uneAbrevJour$ = Mid(unPar.monJourSemaine, 1, 1) + LCase(Mid(unPar.monJourSemaine, 2, 1))
            ImprimerTexte uneMargeG + UnCmEnTwips * 3.1, unCurY, uneAbrevJour$, unNbLignes
            ImprimerTexte uneMargeG + UnCmEnTwips * 3.45, unCurY, Format(unPar.maDate, "dd/mm/yy"), unNbLignes, False
            ImprimerTexte uneMargeG + UnCmEnTwips * 4.47, unCurY, Mid(Format(unPar.monHeureDebut), 1, 5), unNbLignes
            ImprimerTexte uneMargeG + UnCmEnTwips * 5.25, unCurY, Format(unPar.maVmoy, "fixed"), unNbLignes
            ImprimerTexte uneMargeG + UnCmEnTwips * 6.75, unCurY, Format(unPar.maVmin, "fixed"), unNbLignes
            ImprimerTexte uneMargeG + UnCmEnTwips * 8.2, unCurY, Format(unPar.maVmax, "fixed"), unNbLignes
            ImprimerTexte uneMargeG + UnCmEnTwips * 9.7, unCurY, Format(unPar.monNbArret), unNbLignes
            ImprimerTexte uneMargeG + UnCmEnTwips * 10.9, unCurY, FormatterTempsEnHMNS(unPar.monTpsArret), unNbLignes
            ImprimerTexte uneMargeG + UnCmEnTwips * 12.55, unCurY, Format(unPar.monNbDbTop), unNbLignes
            ImprimerTexte uneMargeG + UnCmEnTwips * 13.95, unCurY, FormatterTempsEnHMNS(unPar.monTpsDbTop), unNbLignes
            ImprimerTexte uneMargeG + UnCmEnTwips * 15.6, unCurY, FormatterTempsEnHMNS(unPar.monTFinSection - unPar.monTDebSection), unNbLignes
            ImprimerTexte uneMargeG + UnCmEnTwips * 17.25, unCurY, Format(CLng(unPar.maDistParSection / 10), "#,###,###"), unNbLignes
            'Dessin de la ligne plancher de la ligne i du tableau
            Printer.DrawStyle = vbSolid
            Printer.Line (uneMargeG - 30, unCurY + unDecH + unDecH0 + 10)-(uneMargeG - 30 + UnCmEnTwips * 18.6, unCurY + unDecH + unDecH0 + 10), 0
            'Dessin de la ligne vertical fermant la colonne
            Printer.Line (uneMargeG - 30 + UnCmEnTwips * 18.6, unCurY - 10)-(uneMargeG - 30 + UnCmEnTwips * 18.6, unCurY + unDecH + unDecH0 + 10), 0
            'Augmentation du Y pour la ligne suivante
            unCurY = unCurY + unDecH + unDecH0 + 20
        End If
    Next i
    
    'Impression de la section de travail et du titre
    Printer.Font.Bold = True
    Printer.CurrentY = unCurY + unDecH
    If monIti.CheckSection.Value = 1 Then
        Printer.CurrentX = uneMargeG
        Printer.Print unTitre + " - " + "Section de travail : Entre les repères " + monIti.ComboRepDebSec.Text + " et " + monIti.ComboRepFinSec.Text
    Else
        Printer.CurrentX = uneMargeG
        Printer.Print unTitre + " - " + "Section de travail : Tout l'itinéraire"
    End If
    Printer.Font.Bold = False
    
    'Modif de la marge du haut de la fenêtre itinéraire
    monIti.maMargeHaut = Printer.CurrentY + unDecH * 3
    
    'Remise en trait noir
    Printer.ForeColor = QBColor(0)
End Sub


Private Sub ImprimerTexte(unX As Single, unY As Single, unText As String, Optional unNbLignes As Byte = 1, Optional unDessinLigne As Boolean = True)
    unDecH = Printer.TextHeight("W")
    Printer.CurrentX = unX
    Printer.CurrentY = unY
    Printer.Print unText
    If unDessinLigne Then Printer.Line (unX - 30, unY - 10)-(unX - 30, unY + unNbLignes * (unDecH + 10)), 0
End Sub


Private Sub RemplirTabModifEchImp()
    'Remplissage du tableau des min/max des échelles d'impression
    'Mise en couleur des info-bulles (jaune) des cellules lockés non modifiables
    SpreadEchImp.LockBackColor = vbInfoBackground
    'Remplissage de la première ligne du tableau
    SpreadEchImp.Row = 1
    SpreadEchImp.Col = 1
    SpreadEchImp.Text = monIti.monMinD
    SpreadEchImp.Col = 2
    SpreadEchImp.Text = monIti.monMinT
    SpreadEchImp.Col = 3
    SpreadEchImp.Text = monIti.monMinV
    'Remplissage de la deuxième ligne du tableau
    SpreadEchImp.Row = 2
    SpreadEchImp.Col = 1
    SpreadEchImp.Text = monIti.monMaxD
    SpreadEchImp.Col = 2
    SpreadEchImp.Text = monIti.monMaxT
    SpreadEchImp.Col = 3
    SpreadEchImp.Text = monIti.monMaxV
    'Remplissage de la troisième ligne du tableau
    SpreadEchImp.Row = 3
    SpreadEchImp.Col = 1
    SpreadEchImp.Text = monMinDImp
    SpreadEchImp.Col = 2
    SpreadEchImp.Text = monMinTImp
    SpreadEchImp.Col = 3
    SpreadEchImp.Text = monMinVImp
    'Remplissage de la troisième ligne du tableau
    SpreadEchImp.Row = 4
    SpreadEchImp.Col = 1
    SpreadEchImp.Text = monMaxDImp
    SpreadEchImp.Col = 2
    SpreadEchImp.Text = monMaxTImp
    SpreadEchImp.Col = 3
    SpreadEchImp.Text = monMaxVImp
End Sub

Private Sub SpreadEchImp_KeyUp(KeyCode As Integer, Shift As Integer)
    'Mise à jour des min/max d'impressions
    'On se met dans la ligne et colonne active pour avoir le spreadechimp.text
    SpreadEchImp.Row = SpreadEchImp.ActiveRow
    SpreadEchImp.Col = SpreadEchImp.ActiveCol
    If SpreadEchImp.ActiveRow = 3 Then
        If SpreadEchImp.ActiveCol = 1 Then
            monMinDImp = SpreadEchImp.Text
        ElseIf SpreadEchImp.ActiveCol = 2 Then
            monMinTImp = SpreadEchImp.Text
        ElseIf SpreadEchImp.ActiveCol = 3 Then
            monMinVImp = SpreadEchImp.Text
        End If
    ElseIf SpreadEchImp.ActiveRow = 4 Then
        If SpreadEchImp.ActiveCol = 1 Then
            monMaxDImp = SpreadEchImp.Text
        ElseIf SpreadEchImp.ActiveCol = 2 Then
            monMaxTImp = SpreadEchImp.Text
        ElseIf SpreadEchImp.ActiveCol = 3 Then
            monMaxVImp = SpreadEchImp.Text
        End If
    End If
End Sub

