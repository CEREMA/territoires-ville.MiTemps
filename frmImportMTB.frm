VERSION 5.00
Begin VB.Form frmImportMTB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   Icon            =   "frmImportMTB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextDatePar 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Height          =   300
      Left            =   6180
      Locked          =   -1  'True
      MaxLength       =   26
      TabIndex        =   25
      Top             =   480
      Width           =   2415
   End
   Begin VB.Frame FrameDebFinIti 
      Caption         =   "Modifier l'itinéraire"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   8760
      TabIndex        =   20
      Top             =   3960
      Width           =   2055
      Begin VB.TextBox TextCoefEta 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Height          =   300
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   27
         Top             =   1200
         Width           =   855
      End
      Begin VB.VScrollBar VScrollDebIti 
         Height          =   1080
         LargeChange     =   1000
         Left            =   180
         SmallChange     =   50
         TabIndex        =   24
         Top             =   480
         Width           =   300
      End
      Begin VB.VScrollBar VScrollFinIti 
         Height          =   1080
         LargeChange     =   1000
         Left            =   1620
         SmallChange     =   100
         TabIndex        =   23
         Top             =   480
         Width           =   300
      End
      Begin VB.Label LabelCoefEta 
         Alignment       =   2  'Center
         Caption         =   "Coefficient d' étirement : "
         Height          =   435
         Left            =   600
         TabIndex        =   28
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label LabelFin 
         AutoSize        =   -1  'True
         Caption         =   "Fin"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1620
         TabIndex        =   22
         Top             =   240
         Width           =   210
      End
      Begin VB.Label LabelDeb 
         AutoSize        =   -1  'True
         Caption         =   "Début"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame FrameLegende 
      Caption         =   "Légende"
      Height          =   780
      Left            =   8760
      TabIndex        =   13
      Top             =   5760
      Width           =   2055
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Repère parcours"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1185
      End
      Begin VB.Shape ShapeRepPar 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   120
         Left            =   1800
         Top             =   480
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Repère itinéraire"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1155
      End
      Begin VB.Shape ShapeRepIti 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         FillColor       =   &H0000FF00&
         Height          =   120
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   300
         Width           =   120
      End
   End
   Begin VB.CommandButton BtnAide 
      Caption         =   "Aide ou F1"
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   3450
      Width           =   2055
   End
   Begin VB.CommandButton BtnChgDebIti 
      Caption         =   "Changer début d'itinéraire"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   2490
      Width           =   2055
   End
   Begin VB.CommandButton BtnChgFinIti 
      Caption         =   "Changer fin d'itinéraire"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   2970
      Width           =   2055
   End
   Begin VB.TextBox TextNomPar 
      Height          =   300
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   9
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox TextFichItiRef 
      BackColor       =   &H80000018&
      Height          =   300
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Aucun"
      Top             =   120
      Width           =   7775
   End
   Begin VB.CommandButton BtnZoomIti 
      Caption         =   "Zoom cadré sur l'itinéraire"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8760
      TabIndex        =   4
      Top             =   1530
      Width           =   2055
   End
   Begin VB.CommandButton BtnImport 
      Caption         =   "Importer le parcours cadré"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   1050
      Width           =   2055
   End
   Begin VB.CommandButton BtnCharger 
      Caption         =   "Charger l'itinéraire..."
      Height          =   375
      Left            =   8760
      TabIndex        =   2
      Top             =   570
      Width           =   2055
   End
   Begin VB.CommandButton BtnZoomTout 
      Caption         =   "Zoom Tout les parcours"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8760
      TabIndex        =   5
      Top             =   2010
      Width           =   2055
   End
   Begin VB.CommandButton BtnFermer 
      Cancel          =   -1  'True
      Caption         =   "Fermer"
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   90
      Width           =   2055
   End
   Begin VB.PictureBox ZoneDessin 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FF0000&
      Height          =   4935
      Left            =   60
      ScaleHeight     =   4875
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   840
      Width           =   8535
      Begin VB.Label LabelFinIti 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fin"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3480
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label LabelDebIti 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Début"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3480
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape ShapeRep 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H0000FF00&
         Height          =   120
         Index           =   0
         Left            =   4560
         Shape           =   3  'Circle
         Top             =   2760
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Line LineRep 
         Index           =   0
         Visible         =   0   'False
         X1              =   1440
         X2              =   4440
         Y1              =   2760
         Y2              =   2760
      End
   End
   Begin VB.Label LabelDatePar 
      AutoSize        =   -1  'True
      Caption         =   "Date de mesure : "
      Height          =   195
      Left            =   4800
      TabIndex        =   26
      Top             =   540
      Width           =   1260
   End
   Begin VB.Label LabelColDebut 
      AutoSize        =   -1  'True
      Caption         =   "Couleur du Repère début itinéraire"
      Height          =   195
      Left            =   600
      TabIndex        =   19
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Shape ShapeRepDeb 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      Height          =   120
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   6060
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label LabelColFin 
      AutoSize        =   -1  'True
      Caption         =   "Couleur du Repère fin itinéraire"
      Height          =   195
      Left            =   840
      TabIndex        =   18
      Top             =   6360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape ShapeRepFin 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      Height          =   120
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   6420
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label LabelNomPar 
      AutoSize        =   -1  'True
      Caption         =   "Parcours à importer : "
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   540
      Width           =   1500
   End
   Begin VB.Label LabelItiRef 
      AutoSize        =   -1  'True
      Caption         =   "Itinéraire : "
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmImportMTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private monNbRepIti As Integer, monSaveNbRepIti As Integer
Private monNbRepTot As Integer, monNbPar As Integer
Private monIndRepItiDeb As Integer, monIndRepItiFin As Integer
Private maMargeD As Single, maMargeB As Single, monDec As Single
Private monFichId As Byte, monNbCarLig1a5 As Long
Private monIndRepItiSel As Integer
Private monIndAncienParChoisi As Integer
'Variable indiquant si les repères début et fin de l'itinéraire chargée
'ont été modifié
Private maModifRepItiDeb As Boolean
Private maModifRepItiFin As Boolean
'Collection stockant les indices des parcours qui auront été importés
Private maColParImport As New Collection
'Variables stockant les min et max des temps et vitesse
'de l'itinéraire de référence chargé
Private monMaxT As Single, monMaxV As Single
Private monMinT As Single, monMinV As Single
'Variables stockant les max et min réels en Y suivant le zoom tout ou iti
Private monMaxYReel As Single, monMinYReel As Single
'Variables stockant les max et min réels en X suivant le zoom tout ou iti
Private monMaxXReel As Single, monMinXReel As Single
'Variables stockant les min et max réels en temps et distances
'en zoom tout et en zoom itinéraire
Private monMaxTReelZTout As Single, monMaxDReelZTout As Single
Private monMinTReelZTout As Single, monMinDReelZTout As Single
Private monMaxTReelZIti As Single, monMaxDReelZIti As Single
Private monMinTReelZIti As Single, monMinDReelZIti As Single
'Variable indiquant si les scrollers verticaux des repères début et fin ont
'été modifiés par programme (.value = qqchose) ou par click sur leurs flèches
Private monIsScrollByClick As Boolean


Private Sub BtnAide_Click()
    's'il n'y pas de fichier d'aide pour le projet, afficher un message à l'utilisateur
    'vous pouvez définir le fichier d'aide de votre application dans la boîte
    'de dialogue de propriétés du projet
    If Len(App.HelpFile) = 0 Then
        MsgBox "Impossible d'afficher le sommaire de l'aide. Il n'y a pas d'aide associée à ce projet.", vbInformation, Me.Caption
    Else
        'Lance l'aide du bon contexte
        frmMain.dlgCommonDialog.HelpCommand = cdlHelpContext
        frmMain.dlgCommonDialog.HelpContext = HelpContextID
        frmMain.dlgCommonDialog.ShowHelp  ' affiche la rubrique
    End If
End Sub

Private Sub BtnCharger_Click()
    'Ouvre l'itinéraire contenue dans le fichier choisi dans cette fonction
    'et affiche ses repères
    Dim uneAbsCurv As Long, unTypeIco As Byte
    Dim unNbCarLig1a5 As Long
    Dim uneString As String, uneString1 As String
    Dim unTabSng(1 To NbPasMax) As Single 'Pour stocker les lectures de single
    Dim unCheckSection As Integer, unIndRepDeb As Integer, unIndRepFin As Integer
    Dim unNomFich As String

    'Si protection invalide on ne fait rien
    'If ProtectCheck(2) <> 0 Then Exit Sub
    
    'Vidage de la collection des indices de parcours importés
    ViderCollection maColParImport
    
    'Fermeture du fichier chargé éventuel
    If monFichId > 0 Then Close #monFichId
    
    'Choix du fichier itinéraire
    unNomFich = frmMain.ChoisirFichier(MsgOpen, MsgMitFile, CurDir)
    If unNomFich = "" Then Exit Sub 'Si aucun fichier choisi ou annuler, on sort
    
    'Affichage du sablier pointeur souris d'attente
    Me.MousePointer = vbHourglass
    
    'Initialisation
    monMaxDReelZIti = -100
    monMinDReelZIti = 10000000
    
    'Afichage du nom du fichier itinéraire
    TextFichItiRef.Text = unNomFich
    'Affichage de la fin du nom de fichier si dépasse zone texte
    'en plaçant le curseur en fin de texte
    TextFichItiRef.SelStart = Len(TextFichItiRef.Text)
        
    If ShapeRep.Count > 1 And BtnZoomTout.Enabled = True Then
        'Cas où on charge un autre fichier itinéraire à la place de celui
        'déjà chargé, donc des shaperep existe déjà sauf au premier chargement d'un
        'fichier itinéraire ==> ReDessin des parcours issus du fichier MTB = campagne
        'de mesures en zoom tout les parcours si on était en zoom cadré sur itinéraire
        ZoneDessin.Cls
        DessinerAxes
        DessinerDesParcours ZoneDessin, maColParcoursMTB, maMargeD, monDec, maMargeB, maMargeB, monMinTReelZTout, monMaxTReelZTout, monMinDReelZTout, monMaxDReelZTout, monIndAncienParChoisi
    End If
    
    'Lecture du fichier .mit
    ' Active la routine de gestion d'erreur.
    'MsgBox "Suppression du On Error GoTo ErreurLecture"
    On Error GoTo ErreurReadIti
    
    'Ouverture du fichier en lecture lockée pour éviter deux ouvertures
    monFichId = FreeFile(0) 'renvoi d'un nombre entre 1 et 255
    Open unNomFich For Input Lock Read Write As #monFichId
        
    'Lecture de l'entête des fichiers *.mit
    Input #monFichId, uneString
    If uneString <> "Fichier " + App.Title Then
        'Cas d'un fichier qui n'est pas un fichier MiTemps
        '===> Fermeture du fichier.
        Close #monFichId
        MsgBox MsgErreur + MsgFileNotFile + App.Title + "version " + App.Major + "." + App.Minor, vbCritical
    Else
        'Cas d'un fichier MiTemps *.Mit de la version 3.0
        '1ère ligne du fichier MIT = "Fichier MiTemps"
        
        'Lecture des libellés des conditions météo
        Input #monFichId, uneString
        
        'Lecture des min et max total en distance (parcours complet sans section)
        Input #monFichId, monMinDReelZIti, monMaxDReelZIti
        'Récupération des min et max en distance, vitesse et temps
        Input #monFichId, unTabSng(1), unTabSng(2), monMinV, monMaxV, monMinT, monMaxT
        
        'Lecture du nom de l'itinéraire et de sa longueur
        Input #monFichId, uneString, uneString1
        
        'Récupération des données de la section de travail
        Input #monFichId, unCheckSection, unIndRepDeb, unIndRepFin
        
        'Stockage de la position de lecture juste avant la lecture du nombre
        'de repères et parcours, on s'en sert lors de l'importation pour écrire
        'dans le fichier mit de l'itinéraire chargé le nombre de parcours + 1
        unNbCarLig1a5 = Seek(monFichId)
        
        'Récupération du nombre de repères et de parcours
        Input #monFichId, monNbRepIti, monNbPar
        'Sauvegarde de l'ancien nombre de rep iti, avant le découpage
        'On s'en sert dans la fonction CouperParcoursEntreD1D2 plus bas
        'pour avoir le bon nombre de repères en Control Shape et éviter des
        'plantages dans la fonction CouperParcoursEntreD1D2
        monSaveNbRepIti = monNbRepIti
        'Récupération et création des repères
        uneAbsMax = -100
        uneAbsMin = 10000000
        For i = 1 To monNbRepIti
            Input #monFichId, uneString, uneString1, uneAbsCurv, unTypeIco
            If i > monNbRepTot Then
                'Cas où il faut créer un contrôle rep de plus
                'sinon on modifie juste ses paramètres
                Load ShapeRep(i)
                Load LineRep(i)
                monNbRepTot = monNbRepTot + 1
            End If
            ShapeRep(i).Visible = True
            ShapeRep(i).BackColor = ShapeRepIti.BackColor 'Couleur du rep intermédiaire
            ShapeRep(i).BorderColor = ShapeRepIti.BorderColor
            ShapeRep(i).Tag = Format(uneAbsCurv)
            LineRep(i).Visible = True
            'Recherche de l'abs curv min et max pour les repères
            'début et fin d'itinéraire
            If uneAbsCurv >= uneAbsMax Then
                uneAbsMax = uneAbsCurv
                monIndRepItiFin = i
            End If
            If uneAbsCurv <= uneAbsMin Then
                uneAbsMin = uneAbsCurv
                monIndRepItiDeb = i
            End If
        Next i
        
        'Dessin des repères itinéraire au zoom tout
        DessinerRepereIti ZoneDessin, maMargeD, monDec, maMargeB, maMargeB, monMinTReelZTout, monMaxTReelZTout, monMinDReelZTout, monMaxDReelZTout
        'Indication du repère début et fin de l'itinéraire chargé
        ShapeRep(monIndRepItiDeb).BackColor = ShapeRepDeb.BackColor
        ShapeRep(monIndRepItiDeb).BorderColor = ShapeRepDeb.BorderColor
        ShapeRep(monIndRepItiFin).BackColor = ShapeRepFin.BackColor
        ShapeRep(monIndRepItiFin).BorderColor = ShapeRepFin.BorderColor
    End If
    
    'Masquage des autres contrôles repères qui ne servent plus
    For i = monNbRepIti + 1 To monNbRepTot
        ShapeRep(i).Visible = False
        LineRep(i).Visible = False
    Next i
    
    'Activation du bouton Zoom iti et des boutons de changment et de fin
    'de l'itinéraire et du parcours et déactivation du bouton zoom tout les parcours
    'car c'est le zoom actuel
    BtnZoomTout.Enabled = False
    BtnZoomIti.Enabled = True
    BtnChgDebIti.Enabled = True
    BtnChgFinIti.Enabled = True
    BtnImport.Enabled = True
    
    'Activation des scrollers permettant la modif des débuts et fin d'itinéraire
    FrameDebFinIti.Enabled = True
    LabelDeb.Enabled = True
    LabelFin.Enabled = True
    VScrollDebIti.Enabled = True
    VScrollFinIti.Enabled = True
    
    'Stockage des min et max réels en X et Y
    monMaxYReel = monMaxDReelZTout
    monMinYReel = monMinDReelZTout
    monMaxXReel = monMaxTReelZTout
    monMinXReel = monMinTReelZTout
    
    'Initialisation de l'index du repère iti sélectionné, 0 = rien de sélectionner
    monIndRepItiSel = 0
    
    'Initialisation du coef d'étalonnage (4 chiffres après la virgule comme dans
    'les fichiers MTB MiTemps) et on dégrise son affichage
    TextCoefEta.Text = Format(1, "###.0000")
    LabelCoefEta.Enabled = True
    TextCoefEta.Enabled = True
    
    'Stockage de la position de lecture juste avant la lecture du nombre
    'de repères et parcours, on s'en sert lors de l'importation pour écrire
    'dans le fichier mit de l'itinéraire chargé le nombre de parcours + 1
    monNbCarLig1a5 = unNbCarLig1a5
        
    'Initialisation de l'état d'une modif du repère début ou/et fin
    maModifRepItiDeb = False
    maModifRepItiFin = False
    'Affichage du pointeur souris par défaut
    Me.MousePointer = vbDefault
    ' Quitte pour éviter le gestionnaire d'erreur et on le désactive.
    On Error GoTo 0
    Exit Sub
    
    ' Routine de gestion d'erreur qui évalue le numéro d'erreur.
ErreurReadIti:
    
    ' Traite les autres situations ici...
    unMsg = MsgOpenError + unNomFich + Chr(13) + Chr(13) + MsgErreur + Format(Err.Number) + " : " + Err.Description
    If Err.Number = 70 Then unMsg = unMsg + " (" + UCase(MsgDejaOpen) + ")"
    MsgBox unMsg, vbCritical
    'fermeture du fichier
    Close #monFichId
    Me.MousePointer = vbDefault
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub BtnChgDebIti_Click()
    'Changement de localistation du repère début d'itinéraire
    If monIndRepItiSel = 0 Then
        'aucun rep iti sélectionné
        MsgBox "Il faut d'abord choisir un repère de l'itinéraire en cliquant dessus.", vbInformation
    ElseIf CSng(ShapeRep(monIndRepItiSel).Tag) >= CSng(ShapeRep(monIndRepItiFin).Tag) Then
        'Cas où Abs curv rep sélectionné pour être repère début est
        ' > abs curv repère fin iti ==> Impossible
        MsgBox "Le repère début d'itinéraire ne peut pas être après le repère fin d'itinéraire.", vbExclamation
    Else
        'Cas où l'on peut modifier le repère début
        'Affichage du nouveau repère début avec sa couleur et son label
        ShapeRep(monIndRepItiSel).BackColor = ShapeRep(monIndRepItiDeb).BackColor
        ShapeRep(monIndRepItiSel).BorderColor = ShapeRep(monIndRepItiDeb).BorderColor
        LabelDebIti.Top = LineRep(monIndRepItiSel).Y1 - LabelDebIti.Height / 2
        'Mise en couleur normale de l'ancien repère début
        ShapeRep(monIndRepItiDeb).BackColor = ShapeRepIti.BackColor
        ShapeRep(monIndRepItiDeb).BorderColor = ShapeRepIti.BorderColor
        'Stockage du nouveau index de repère début
        monIndRepItiDeb = monIndRepItiSel
        'Mise à jour du min pour le zoom cadré sur l'itinéraire
        monMinDReelZIti = CSng(ShapeRep(monIndRepItiSel).Tag)
        
        'Synchronisation de la nouvelle position dans le vscroll début
        'Indication d'une modif des vscrollers repères début et fin par programme
        '(.value = qqchose)==> le change event sur vscroll ne fera rien
        monIsScrollByClick = False
        'Positionnement du vscroll pour le repère début par rapport
        'au Y écran du repère début d'itinéraire
        VScrollDebIti.Value = ConvertirYEcranEnVScrollValue(VScrollDebIti, LineRep(monIndRepItiSel).Y1)
        'Remise de l'Indication qu'une modif des vscrollers repères début et fin
        'par click souris sur leurs flèches est possible
        '==> le change event sur vscroll fera son travail dans ces cas
        monIsScrollByClick = True
        
        'Indication d'une modif du repère début
        maModifRepItiDeb = True
    End If
End Sub

Private Sub BtnChgFinIti_Click()
    'Changement de localistation du repère fin d'itinéraire
    If monIndRepItiSel = 0 Then
        'aucun rep iti sélectionné
        MsgBox "Il faut d'abord choisir un repère de l'itinéraire en cliquant dessus.", vbInformation
    ElseIf CSng(ShapeRep(monIndRepItiSel).Tag) <= CSng(ShapeRep(monIndRepItiDeb).Tag) Then
        'Cas où Abs curv rep sélectionné pour être repère fin est
        ' < abs curv repère début iti ==> Impossible
        MsgBox "Le repère fin d'itinéraire ne peut pas être avant le repère début d'itinéraire.", vbExclamation
    Else
        'Cas où l'on peut modifier le repère fin
        'Affichage du nouveau repère fin avec sa couleur et son label
        ShapeRep(monIndRepItiSel).BackColor = ShapeRep(monIndRepItiFin).BackColor
        ShapeRep(monIndRepItiSel).BorderColor = ShapeRep(monIndRepItiFin).BorderColor
        LabelFinIti.Top = LineRep(monIndRepItiSel).Y1 - LabelFinIti.Height / 2
        'Mise en couleur normale de l'ancien repère fin
        ShapeRep(monIndRepItiFin).BackColor = ShapeRepIti.BackColor
        ShapeRep(monIndRepItiFin).BorderColor = ShapeRepIti.BorderColor
        'Stockage du nouveau index de repère fin
        monIndRepItiFin = monIndRepItiSel
        'Mise à jour du max pour le zoom cadré sur l'itinéraire
        monMaxDReelZIti = CSng(ShapeRep(monIndRepItiSel).Tag)
        
        'Synchronisation de la nouvelle position dans le vscroll fin
        'Indication d'une modif des vscrollers repères début et fin par programme
        '(.value = qqchose)==> le change event sur vscroll ne fera rien
        monIsScrollByClick = False
        'Positionnement du vscroll pour le repère fin par rapport
        'au Y écran du repère fin d'itinéraire
        VScrollFinIti.Value = ConvertirYEcranEnVScrollValue(VScrollDebIti, LineRep(monIndRepItiSel).Y1)
        'Remise de l'Indication qu'une modif des vscrollers repères début et fin
        'par click souris sur leurs flèches est possible
        '==> le change event sur vscroll fera son travail dans ces cas
        monIsScrollByClick = True
        
        'Indication d'une modif du repère fin
        maModifRepItiFin = True
    End If
End Sub


Private Sub btnFermer_Click()
    Unload Me
End Sub

Private Sub BtnImport_Click()
    Dim unParDejaImport As Integer, unCoefEtirement As Single
    Dim unTextLine As String, unTextLine1 As String, unTextLine2 As String
    Dim unParToImport As Parcours, unMsg As String
    Dim uneD1 As Single, uneD2 As Single, uneD0 As Single
    Dim unNbRepIti As Integer, unNbPar As Integer
    Dim unNomFich As String, unNomFich0 As String
    Dim unNbRepTop As Integer, uneAbsCurv As Long, unOldFichId As Byte
    Dim unNomLong As String, unNomCourt As String, unTypeIco As Integer
    Dim uneColString As New Collection, unNomIti As String, uneLongIti As String
    Dim unIndRepItiDeb As Integer, unIndRepItiFin As Integer
    Dim unMinD As Single, unMaxD As Single
    Dim unMinV As Single, unMaxV As Single
    Dim unMinT As Single, unMaxT As Single, uneDuree As Single
    Dim unI1 As Integer, unI2 As Integer, unI3 As Integer
    Dim unR1 As Single, unR2 As Single, unR3 As Single
    Dim unR4 As Single, unR5 As Single, unR6 As Single
        
    If monIndAncienParChoisi = 0 Then
        'aucun parcours sélectionné
        MsgBox "Il faut d'abord choisir un parcours en cliquant sur sa courbe Distance / Temps.", vbInformation
    Else
        'Récupération du parcours à importer et des abscisses curvilignes
        'des repères début et fin de l'itinéraire chargé
        Set unParToImport = maColParcoursMTB(monIndAncienParChoisi)
        'Récupération des abs curv de début et de fin
        uneD1 = CSng(Format(ShapeRep(monIndRepItiDeb).Tag))
        uneD2 = CSng(Format(ShapeRep(monIndRepItiFin).Tag))
        'On fait en sorte que D1 <= D2
        If uneD1 > uneD2 Then
            uneD0 = uneD1
            uneD1 = uneD2
            uneD2 = uneD0
        End If
        'Test si les repères début ou/et fin de l'itinéraire ont été changé
        'On invite l'utilisateur à choisir un nouveau fichier d'itinéraire qui ne
        'contiendra aucun parcours et uniquement les repères entre les nouveaux
        'début et fin et qui servira à stocker les parcours importés par coupure
        'entre les nouveaux repères début et fin
        If maModifRepItiDeb Or maModifRepItiFin Then
            'Message d'avertissement et de confirmation de continuation d'action
            unMsg = "Les repères début et/ou fin de l'itinéraire de référence ont été changés." + Chr(13) + Chr(13)
            unMsg = unMsg + "Pour éviter d'avoir des parcours avec un nombre de repères différents," + Chr(13)
            unMsg = unMsg + "vous allez devoir choisir un nouveau fichier itinéraire qui ne contiendra" + Chr(13)
            unMsg = unMsg + "aucun parcours et uniquement les repères entre les nouveaux repères" + Chr(13)
            unMsg = unMsg + "début et fin et qui servira à stocker les parcours qui seront importés" + Chr(13)
            unMsg = unMsg + "par coupure entre ces nouveaux repères début et fin."
            unMsg = unMsg + Chr(13) + Chr(13) + "Voulez-vous continuer ?"
            If MsgBox(unMsg, vbYesNo + vbQuestion) = vbNo Then Exit Sub
            
            'Stockage de l'ancien fichier itinéraire chargé
            unNomFich0 = TextFichItiRef.Text
            'Demande du nouveau fichier MIT de stockage
            unNomFich = frmMain.ChoisirFichier(MsgSaveAs, MsgMitFile, CurDir)
            If unNomFich = "" Then Exit Sub 'Si aucun fichier choisi ou annuler, on sort
            
            'Fermeture et réouverture en mode verrouillé pour
            'éviter deux ouvertures de l'ancien fichier itinéraire
            'ainsi sa lecture repart de la 1ère ligne
            unOldFichId = FreeFile(0)
            Close #monFichId
            Open unNomFich0 For Input Lock Read Write As #unOldFichId
            'Lecture des lignes jusqu'à la fin des repères et stockage des deux
            'premières lignes pour insertion dans le new MIT
            Input #unOldFichId, unTextLine1
            Input #unOldFichId, unTextLine2
            Input #unOldFichId, unR1, unR2
            Input #unOldFichId, unR1, unR2, unR3, unR4, unR5, unR6
            'Récupération du nom et la longueur de l'itinéraire
            Input #unOldFichId, unNomIti, uneLongIti
            Input #unOldFichId, unI1, unI2, unI3
            Input #unOldFichId, unI1, unI2
            'Récupération des repères se trouvant entre les repères début et fin et
            'du nombre de parcours que l'on met à 0 et calcul de ce nombres de repères
            unNbPar = 0
            unNbRepIti = 0 'Initialisation pour le calcul du nombre de repères
            For j = 1 To monNbRepIti
                Input #unOldFichId, unNomLong, unNomCourt, uneAbsCurv, unTypeIco
                If uneAbsCurv >= uneD1 - mesOptions.monEcartMax And uneAbsCurv <= uneD2 + mesOptions.monEcartMax Then
                    'On incrémente le nombre de repères total
                    unNbRepIti = unNbRepIti + 1
                    'Calcul des nouveaux indices des repères début et fin
                    If Abs(uneAbsCurv - uneD1) < mesOptions.monEcartMax Then
                        unIndRepItiDeb = unNbRepIti
                    ElseIf Abs(uneAbsCurv - uneD2) < mesOptions.monEcartMax Then
                        unIndRepItiFin = unNbRepIti
                    End If
                    'On ajoute en fin de liste à chaque fois
                    If uneColString.Count = 0 Then
                        uneColString.Add unNomLong
                    Else
                        uneColString.Add unNomLong, , , uneColString.Count
                    End If
                    uneColString.Add unNomCourt, , , uneColString.Count
                    uneColString.Add Format(uneAbsCurv - uneD1), , , uneColString.Count
                    uneColString.Add Format(unTypeIco), , , uneColString.Count
                End If
            Next j
            'Fermeture du fichier itinéraire précédemment chargé
            Close #unOldFichId
            'Ouverture en mode output du nouveau fichier itinéraire pour écrire dedans
            Open unNomFich For Output As #monFichId
            'Ecriture des deux premières lignes du mit de départ
            Write #monFichId, unTextLine1
            Write #monFichId, unTextLine2
            'Ecriture du min et du max total en distance
            Write #monFichId, 0, CLng(uneD2 - uneD1)
            'Calcul des min et max en distance, vitesse et durée
            'en convertissant D1 et D2 des mètres en décimètres
            'Valeur initiale 1 km/h pour V et 1 minutes pour T, 0 pour les min
            unMaxV = 1
            unMaxT = 1
            unMaxD = uneD2 - uneD1
            unMinV = 1000
            unMinT = 0
            unMinD = 0
            unParToImport.CalculerLesVitDistDureeEtArrets CLng(uneD1 * 10), CLng(uneD2 * 10)
            If unParToImport.maVmax > unMaxV Then unMaxV = unParToImport.maVmax
            If unParToImport.maVmin < unMinV Then unMinV = unParToImport.maVmin
            uneDuree = unParToImport.monTFinSection - unParToImport.monTDebSection
            'Conversion du temps des dixièmes de secondes en minutes
            uneDuree = uneDuree / 600
            If unMaxT < uneDuree Then unMaxT = uneDuree
            'Ecriture des min et max en distance, vitesse et durée
            Write #monFichId, unMinD, CLng(unMaxD), unMinV, CLng(unMaxV), unMinT, CLng(unMaxT * 10) / 10
            'Ecriture du nom et de la longueur du nouvel itinéraire de référence
            'Write #monFichId, unNomIti, Format(unMaxD)
            uneStrTmp = unNomIti + " importé entre " + Format(uneD1) + " et " + Format(uneD2) + " mètres"
            Write #monFichId, uneStrTmp, Format(unMaxD)
            'Ecriture de la section de travail, pas de section mais début et fin
            'correspondant aux repères début et fin du nouvel itinéraire donc
            'aux nouveaux indices calculés ci-dessus car les repères ne sont pas
            'classées par ordre croissant
            uneSectionDefinie = 0
            Write #monFichId, uneSectionDefinie, unIndRepItiDeb, unIndRepItiFin
            'Stockage de la position de lecture juste avant la lecture du nombre
            'de repères et parcours, on s'en sert lors de l'importation pour écrire
            'dans le fichier mit de l'itinéraire chargé le nombre de parcours + 1
            monNbCarLig1a5 = Seek(monFichId)
            'Ecriture du nb de repères et de parcours et des repères qui ont été
            'trouvés entre les repères début et fin
            Print #monFichId, Format(unNbRepIti) + "," + Format(unNbPar, "000#")
            For j = 0 To unNbRepIti - 1
                Write #monFichId, uneColString(4 * j + 1), uneColString(4 * j + 2), CLng(uneColString(4 * j + 3)), CInt(uneColString(4 * j + 4))
            Next j
            'On supprime la collection
            ViderCollection uneColString
            Set uneColString = Nothing
                        
            'Fermeture et réouverture en mode verrouillé pour
            'éviter deux ouvertures du nouveau fichier itinéraire
            Close #monFichId
            Open unNomFich For Input Lock Read Write As #monFichId
            'Affichage en haut du nouveau fichier d'itinéraire de référence
            TextFichItiRef.Text = unNomFich
            'Initialisation de l'état d'une modif du repère début ou/et fin
            maModifRepItiDeb = False
            maModifRepItiFin = False
            'Vidage de la collection des parcours déjà importé
            ViderCollection maColParImport
            'Mise à jour du nouveau nombre de parcours et de repères
            monNbRepIti = unNbRepIti
            monNbPar = unNbPar
        End If
        
        'Test de la présence de l'indice du parcours choisi dans la
        'collection des indices des parcours importés
        '0n prévient de l'import mulitple du même parcours pendant la session
        'de travail, mais cela peut servir si on découpe par rapport à des lignes
        'de bus, d'où on laisse l'utilisateur choisir
        unParDejaImport = vbYes
        For i = 1 To maColParImport.Count
            If monIndAncienParChoisi = maColParImport(i) Then
                unMsg = "Le parcours " + maColParcoursMTB(monIndAncienParChoisi).monNom + " mesuré le "
                unMsg = unMsg + maColParcoursMTB(monIndAncienParChoisi).monJourSemaine + " " + Format(maColParcoursMTB(monIndAncienParChoisi).maDate) + " à " + Format(maColParcoursMTB(monIndAncienParChoisi).monHeureDebut)
                unMsg = unMsg + " a déjà été importé dans le fichier " + TextFichItiRef.Text
                unMsg = unMsg + Chr(13) + Chr(13) + "Voulez-vous continuer ?"
                unParDejaImport = MsgBox(unMsg, vbYesNo + vbQuestion)
                'unParDejaImport = True
                Exit For
            End If
        Next i
        'If unParDejaImport = False Then
        If unParDejaImport = vbYes Then
            unMsg = "Voulez-vous importer le parcours " + maColParcoursMTB(monIndAncienParChoisi).monNom + " mesuré le "
            unMsg = unMsg + maColParcoursMTB(monIndAncienParChoisi).monJourSemaine + " " + Format(maColParcoursMTB(monIndAncienParChoisi).maDate) + " à " + Format(maColParcoursMTB(monIndAncienParChoisi).monHeureDebut)
            unMsg = unMsg + " en le nommant " + Chr(34) + TextNomPar.Text + Chr(34) + " dans le fichier " + TextFichItiRef.Text + " ?"
            If MsgBox(unMsg, vbQuestion + vbYesNo) = vbNo Then
                'Sortie directe si on ne veut pas importer le parcours sélectionné
                Exit Sub
            End If
                                    
            'Coupure du parcours à importer entre les repères début et fin
            'de l'itinéraire chargé et on met les données dans le parcours
            'donnée par la variable globale monParToImport
            unCoefEtirement = Format(CSng(TextCoefEta.Text))
            If CouperParcoursEntreD1D2(unParToImport, Me, monSaveNbRepIti, monIndRepItiDeb, monIndRepItiFin, unCoefEtirement) = False Then
                'Si la coupure ne peut se faire, on sort sans rien faire
                Exit Sub
            End If
            
            ' Active la routine de gestion d'erreur.
            'MsgBox "Suppression du On Error GoTo ErreurRWIti"
            On Error GoTo ErreurRWIti
            
            'Fermeture du fichier itinéraire chargé pour réouverture en append après
            Close #monFichId
            
            'Ouverture du fichier itinéraire chargé en mode ajout pour ajouter les
            'données du parcours importé
            unNomFich = TextFichItiRef.Text
            Open unNomFich For Append As #monFichId
            monParToImport.monNom = TextNomPar.Text
            EcrireDonneesParcoursDansFichierMIT monFichId, monParToImport
            
            'On se met sur la ligne où se trouve le nombre de repères et de parcours
            'pour ajouter +1 au nombre de parcours en le formattant sur 4 caractères
            'pour écrire juste dans sa place (4 caractères lors de la sauvegarde du
            'fichier MIT).
            'Cette position est donnée par monNbCarLig1a5 définie dans la fonction
            'BtnCharger_Click liant le fichier itinéraire chargé ou dans cette
            'procédure plus haut dans la partie changeant le fichier en cas de modif
            'de repère début et/ou fin
            Seek #monFichId, monNbCarLig1a5
            unTextLine = Format(monNbRepIti) + "," + Format(monNbPar + 1, "000#")
            Print #monFichId, unTextLine
            
            'Fermeture et réouverture en mode verrouillé pour
            'éviter deux ouvertures
            Close #monFichId
            Open unNomFich For Input Lock Read Write As #monFichId
            'Ajout dans la collection des indices des parcours importés
            maColParImport.Add monIndAncienParChoisi
            'Incrémentation du nombre de parcours total car ici l'importation
            'est finie et c'est bien passée, donc on peut faire +1
            monNbPar = monNbPar + 1
            MsgBox "Importation réussie", vbInformation
        End If
    End If
    
    'Sortie pour éviter la gestion d'erreur
    On Error GoTo 0
    Exit Sub
    
    ' Routine de gestion d'erreur qui évalue le numéro d'erreur.
ErreurRWIti:
    
    ' Traite les autres situations ici...
    unMsg = MsgOpenError + unNomFich + Chr(13) + Chr(13) + MsgErreur + Format(Err.Number) + " : " + Err.Description
    If Err.Number = 70 Then unMsg = unMsg + " (" + UCase(MsgDejaOpen) + ")"
    MsgBox unMsg, vbCritical
    'Fermeture et réouverture en mode verrouillé
    Close #monFichId
    Open unNomFich For Input Lock Read Write As #monFichId
    Me.MousePointer = vbDefault
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub BtnZoomIti_Click()
    Dim unT1 As Single, unT2 As Single
    
    'BtnZoomIti.Enabled = False
    'Dessin des axes Distances en Y et temps en X
    DessinerAxes
    'Calcul des temps de passage aux distances correspondant aux repères
    'début et fin de l'itinéraire chargé
    'Initialisation
    monMaxTReelZIti = -100
    monMinTReelZIti = 10000000
    For i = 1 To maColParcoursMTB.Count
        If maColParcoursMTB(i).monIsUtil Then
            maColParcoursMTB(i).DonnerTemps monMinDReelZIti * 10, monMaxDReelZIti * 10, unT1, unT2
            'conversion des temps des dixièmes de secondes en minutes
            unT1 = unT1 / 600
            unT2 = unT2 / 600
            If unT1 <= monMinTReelZIti Then monMinTReelZIti = unT1
            If unT2 >= monMaxTReelZIti Then monMaxTReelZIti = unT2
        End If
    Next i
    'Dessin des parcours issus du fichier MTB = campagne de mesures
    DessinerDesParcours ZoneDessin, maColParcoursMTB, maMargeD, monDec, maMargeB, maMargeB, monMinTReelZIti, monMaxTReelZIti, monMinDReelZIti, monMaxDReelZIti, monIndAncienParChoisi
    'Dessin des repères itinéraire au zoom tout
    DessinerRepereIti ZoneDessin, maMargeD, monDec, maMargeB, maMargeB, monMinTReelZIti, monMaxTReelZIti, monMinDReelZIti, monMaxDReelZIti
    BtnZoomTout.Enabled = True
    'Stockage des min et max réels en X et Y
    monMaxYReel = monMaxDReelZIti
    monMinYReel = monMinDReelZIti
    monMaxXReel = monMaxTReelZIti
    monMinXReel = monMinTReelZIti
End Sub

Private Sub BtnZoomTout_Click()
    MousePointer = vbHourglass
    
    'Dessin en zoom tout
    DessinerZoomTout
    'Désactivation du bouton Zoom tout
    BtnZoomTout.Enabled = False
    'Activation du bouton Zoom iti
    BtnZoomIti.Enabled = True
    'Stockage des min et max réels en X et Y
    monMaxYReel = monMaxDReelZTout
    monMinYReel = monMinDReelZTout
    monMaxXReel = monMaxTReelZTout
    monMinXReel = monMinTReelZTout
    
    MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    'Pour ne dessiner qu'au premier activate
    If monMaxTReelZTout > 0 And monMaxDReelZTout > 0 Then
        'Après il y a chargement d'un fichier MTB et les deux max sont tjs > 0
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    monIndAncienParChoisi = 0
    Caption = "Import de " + Tag
    
    'Mettre tous les parcours issus du fichier MTB à utiliser pour qu'ils s'affichent
    'For i = 1 To maColParcoursMTB.Count
    '    maColParcoursMTB(i).monIsUtil = True
    'Next i
    
    'Calcul des maxi réelles en distances et en temps pour le zoom total
    DonnerMaxDistDuree maColParcoursMTB, monMaxDReelZTout, monMaxTReelZTout
    'Calcul des mini réelles en distances et en temps pour le zoom total
    monMinDReelZTout = 0
    monMinTReelZTout = 0
    
    'Stockage des min et max réels en X et Y
    monMaxYReel = monMaxDReelZTout
    monMinYReel = monMinDReelZTout
    monMaxXReel = monMaxTReelZTout
    monMinXReel = monMinTReelZTout
    
    'Dessin en zoom tout
    DessinerZoomTout
    
    MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim unControl As Control
    
    'Contexte d'aide
    HelpContextID = HelpID_WinImportMesure
    'Vidage de la collection des indices de parcours importés
    ViderCollection maColParImport
    'Initialisation du nombre de repères créés
    'et du nombre de parcours importés
    monNbRepTot = 0
    monNbRepIti = 0
    monSaveNbRepIti = 0
    monFichId = 0
    'Autres initialisations
    monIndRepItiDeb = 0
    monIndRepItiFin = 0
    monMaxTReelZTout = 0
    monMaxDReelZTout = 0
    monIndAncienParChoisi = 0
    
    'Placement et taille de la fenêtre
    Width = Screen.Width * 0.95
    Height = Screen.Height * 0.9
    CentrerFenetreEcran Me
    
    'Mise contre le bord droit des boutons de commandes
    'et des labels rouge d'info IHM
    For i = 0 To Controls.Count - 1
        Set unControl = Controls(i)
        If TypeOf unControl Is CommandButton Then
            unControl.Left = Width - unControl.Width - 120
        End If
    Next i
        
    'Agrandissement de la zone de dessin
    ZoneDessin.Width = Width - BtnFermer.Width - 240
    ZoneDessin.Height = ScaleHeight - ZoneDessin.Top - 60
    
    'Positionnnment et Agrandissement des textes affichant le nom du
    'fichier itinéraire, le nom du parcours à importer
    'et sa date de mesure
    TextFichItiRef.Width = BtnFermer.Left - TextFichItiRef.Left - 75
    
    TextDatePar.Left = BtnFermer.Left - TextDatePar.Width - 75
    LabelDatePar.Left = TextDatePar.Left - LabelDatePar.Width '- 60
    TextNomPar.Width = LabelDatePar.Left - TextNomPar.Left - 60
    TextNomPar.Text = "Aucun"
    
    'Placement de la légende en bas à droite
    FrameLegende.Top = ScaleHeight - FrameLegende.Height - 60
    FrameLegende.Left = BtnFermer.Left
    
    'Positionnement de la frame permettant la modif des repères début
    'et fin d'un itinéraire
    FrameDebFinIti.Left = BtnFermer.Left
    FrameDebFinIti.Top = BtnAide.Top + BtnAide.Height + 120
    FrameDebFinIti.Height = FrameLegende.Top - FrameDebFinIti.Top - 120
    VScrollDebIti.Height = FrameDebFinIti.Height - 600
    VScrollFinIti.Height = FrameDebFinIti.Height - 600
    VScrollDebIti.Enabled = False
    VScrollFinIti.Enabled = False
    LabelCoefEta.Enabled = False
    TextCoefEta.Enabled = False
    'Indication d'une modif des vscrollers repères début et fin
    'par click souris sur ses flèches ==> DéplacerTousRepIti fera son travail
    monIsScrollByClick = True
End Sub

Public Sub DessinerZoomTout()
    'Dessiner dans la zone de dessin à un niveau de zoom montrant toutes
    'les entités graphiques
    
    'Dessin des axes
    DessinerAxes
    'Dessin des parcours issus du fichier MTB = campagne de mesures
    DessinerDesParcours ZoneDessin, maColParcoursMTB, maMargeD, monDec, maMargeB, maMargeB, monMinTReelZTout, monMaxTReelZTout, monMinDReelZTout, monMaxDReelZTout, monIndAncienParChoisi
    'Dessin des repères itinéraire au zoom tout
    DessinerRepereIti ZoneDessin, maMargeD, monDec, maMargeB, maMargeB, monMinTReelZTout, monMaxTReelZTout, monMinDReelZTout, monMaxDReelZTout
End Sub

Public Sub DessinerRepereIti(uneZoneDes As PictureBox, uneMargeD As Single, uneMargeG As Single, uneMargeB As Single, uneMargeH As Single, unMinXreel As Single, unMaxXreel As Single, unMinYreel As Single, unMaxYreel As Single)
    'Dessin des repères de l'itinéraire de référence chargé dans la large droite d'
    'une picture box avec respect des marges entre un min en X et un max en X réel
    'et entre un min en Y réel et un max Y réel
    Dim unX1 As Single, unX2 As Single, unD1 As Single, unD2 As Single
    Dim unXecran As Single, unYecran As Single
    Dim unXecSuiv As Single, unYecSuiv As Single
    Dim unMinXecran As Single, unMaxXecran As Single
    Dim unMinYecran As Single, unMaxYecran As Single
    Dim uneDistMaxReelX As Single, uneDistMaxEcranX As Single
    Dim uneDistMaxReelY As Single, uneDistMaxEcranY As Single
    
    'Détermination des min/max écran
    'Variables servant pour la conversion coordonnées réelles en écran
    uneDistMaxReelX = unMaxXreel - unMinXreel
    uneDistMaxEcranX = uneZoneDes.Width - uneMargeG - uneMargeD
    uneDistMaxReelY = unMaxYreel - unMinYreel
    uneDistMaxEcranY = uneZoneDes.Height - uneMargeH - uneMargeB

    'Conversion en coordonnées Y écran des distances réelles
    'Les Y sont orientés vers le bas, donc le max réel correspondant au max écran
    'est inférieur au min écran correspondant au min réel
    unMaxYecran = uneMargeH
    unMinYecran = uneZoneDes.Height - uneMargeB
    
    'Conversion en coordonnées X écran des temps ou vitesses réels
    unMinXecran = uneMargeG
    unMaxXecran = uneZoneDes.Width - uneMargeD
    
    'Calcul de X ecran de positionnement du rond symbolisant le repère
    'et la fin de la ligne de rappel
    unXecran = unMaxXecran + monDec
    
    'Dessiner des repères et de leur ligne de rappel
    For i = 1 To monSaveNbRepIti
        'Récup de l'abs curviligne en mètres
        unD1 = CSng(ShapeRep(i).Tag)
        
        'Placement en Y écran du rond symbolisant le repère et de la ligne de rappel
        'Conversion en coordonnées écrans des coordonnées réelles
        unYecran = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD1, uneDistMaxReelY, uneDistMaxEcranY)
        'Dessin du repère et de sa ligne de rappel
        's'il est entre le min et le max y écran
        'min y écran  > max y écran car les y écran orientés vers le bas en Y,
        'donc aprés conversion donnée réelle en écran le max devient < au min
        uneVisu = (unYecran <= unMinYecran + EpsilonEcran And unYecran >= unMaxYecran - EpsilonEcran)
        ShapeRep(i).Visible = uneVisu
        ShapeRep(i).Left = unXecran
        ShapeRep(i).Top = unYecran - TailleRep
        
        LineRep(i).Visible = uneVisu
        LineRep(i).X1 = unMinXecran
        LineRep(i).X2 = unXecran
        LineRep(i).Y1 = unYecran
        LineRep(i).Y2 = unYecran
        
        'Indication du début ou fin d'itinéraire
        If i = monIndRepItiDeb Then
            'Cas du repère Début d'itinéraire
            LabelDebIti.Visible = uneVisu
            LabelDebIti.Left = unXecran + ShapeRep(i).Width
            LabelDebIti.Top = unYecran - LabelDebIti.Height / 2
            VScrollDebIti.Tag = 0
            'Indication d'une modif des vscrollers repères début et fin par programme
            '(.value = qqchose) ==> le change event sur vscroll ne fera rien
            monIsScrollByClick = False
            'Positionnement du vscroll pour le repère début par rapport
            'au Y écran du repère début d'itinéraire
            VScrollDebIti.Value = ConvertirYEcranEnVScrollValue(VScrollDebIti, unYecran)
            'Remise de l'Indication qu'une modif des vscrollers repères début et fin
            'par click souris sur leurs flèches est possible
            '==> le change event sur vscroll fera son travail dans ces cas
            monIsScrollByClick = True
        ElseIf i = monIndRepItiFin Then
            'Cas du repère Fin d'itinéraire
            LabelFinIti.Visible = uneVisu
            LabelFinIti.Left = unXecran + ShapeRep(i).Width
            LabelFinIti.Top = unYecran - LabelFinIti.Height / 2
            'Indication d'une modif des vscrollers repères début et fin par programme
            '(.value = qqchose) ==> le change event sur vscroll ne fera rien
            monIsScrollByClick = False
            'Positionnement du vscroll pour le repère fin par rapport
            'au Y écran du repère fin d'itinéraire
            VScrollFinIti.Value = ConvertirYEcranEnVScrollValue(VScrollFinIti, unYecran)
            'Remise de l'Indication qu'une modif des vscrollers repères début et fin
            'par click souris sur leurs flèches est possible
            '==> le change event sur vscroll fera son travail dans ces cas
            monIsScrollByClick = True
            'On grise le vscroll de modif du repère fin iti s'il n'est pas visible
            VScrollFinIti.Enabled = uneVisu
            LabelFin.Enabled = uneVisu
        End If
    Next i
End Sub

Private Sub DessinerAxes()
    'Procédure dessinant les axes Ox (les temps de parcours)
    'et Oy (les distances parcourues)
    
    monDec = 60
    maMargeB = ZoneDessin.TextHeight("Temps en minutes") + monDec * 2
    maMargeD = 850
    
    'On vide la zone de dessin
    ZoneDessin.Cls
    
    'Axe des temps = Ox
    ZoneDessin.Line (monDec, ZoneDessin.Height - maMargeB)-(ZoneDessin.Width - maMargeD, ZoneDessin.Height - maMargeB), QBColor(0)
    'Dessin de la flèche
    ZoneDessin.Line (ZoneDessin.Width - maMargeD * 2, ZoneDessin.Height - maMargeB)-(ZoneDessin.Width - maMargeD * 2 - monDec, ZoneDessin.Height - maMargeB - monDec), QBColor(0)
    ZoneDessin.Line (ZoneDessin.Width - maMargeD * 2, ZoneDessin.Height - maMargeB)-(ZoneDessin.Width - maMargeD * 2 - monDec, ZoneDessin.Height - maMargeB + monDec), QBColor(0)
    ZoneDessin.CurrentX = ZoneDessin.Width - maMargeD - monDec - ZoneDessin.TextWidth("Temps en minutes")
    ZoneDessin.CurrentY = ZoneDessin.Height - maMargeB + monDec
    ZoneDessin.Print "Temps en minutes"
    
    'Axe des distances = Oy
    ZoneDessin.Line (ZoneDessin.Width - maMargeD, monDec)-(ZoneDessin.Width - maMargeD, ZoneDessin.Height - monDec * 2), QBColor(0)
    ZoneDessin.CurrentX = ZoneDessin.Width - maMargeD - monDec * 2 - ZoneDessin.TextWidth("Distance (m)")
    ZoneDessin.CurrentY = 0 'monDec
    ZoneDessin.Print "Distance (m)"
    'Dessin de la flèche
    ZoneDessin.Line (ZoneDessin.Width - maMargeD, monDec)-(ZoneDessin.Width - maMargeD - monDec, monDec * 2), QBColor(0)
    ZoneDessin.Line (ZoneDessin.Width - maMargeD, monDec)-(ZoneDessin.Width - maMargeD + monDec, monDec * 2), QBColor(0)

    'Affichage du texte "Parcours mesurés" en haut et au milieu de la partie
    'visualisant les parcours issues de la campagne de mesure
    ZoneDessin.CurrentX = (ZoneDessin.Width - maMargeD - monDec * 2 - ZoneDessin.TextWidth("Parcours mesurés")) / 2
    ZoneDessin.CurrentY = 0
    ZoneDessin.Print "Parcours mesurés"
    
    'Affichage du texte "Itinéraire" en haut et au milieu de la partie
    'visualisant les repères de l'itinéraire de référence chargé
    ZoneDessin.CurrentX = ZoneDessin.Width - maMargeD + monDec * 2
    ZoneDessin.CurrentY = 0
    ZoneDessin.Print "Itinéraire"
    
    'Dessin du cadre de dessin possible pour tester
    'ZoneDessin.Line (monDec, maMargeB)-(ZoneDessin.Width - maMargeD, ZoneDessin.Height - maMargeB), QBColor(0), B
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Fermeture du fichier itinéraire chargé éventuel
    If monFichId > 0 Then Close #monFichId
    'Fermeture de la fenêtre de choix lors d'une sélection de plusieurs parcours
    'possibles
    Unload frmChoixPar
End Sub


Private Sub LabelDebIti_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Sélection du repère début d'iti impossible
    MsgBox "La sélection de repère début d'itinéraire est impossible.", vbExclamation
End Sub

Private Sub LabelFinIti_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Sélection du repère fin d'iti impossible
    MsgBox "La sélection de repère fin d'itinéraire est impossible.", vbExclamation
End Sub

Private Sub VScrollDebIti_Change()
    DeplacerTousRepIti
End Sub

Private Sub VScrollDebIti_Scroll()
    DeplacerTousRepIti
End Sub

Private Sub VScrollFinIti_Change()
    EtirerTousRepIti
End Sub

Private Sub VScrollFinIti_Scroll()
    EtirerTousRepIti
End Sub

Private Function ConvertirVScrollValueEnYEcran(unVScroll As VScrollBar, uneVal As Single) As Double
    'Fonction retournant la coordonnées en Y écran donnée par le champ
    'value des scroller vertical liés au début et fin d'itinéraire chargé
    Dim unMaxYecran As Double, unMinYecran As Double
    Dim unA As Double, unB As Double
    
    'Le max y écran correpondant au max y réel est en haut de fenêtre
    'c'est l'inverse pour le min y écran car Oy réel est orienté de bas en haut
    'alors que Oy écran est orienté de haut en bas
    unMaxYecran = maMargeB
    unMinYecran = ZoneDessin.Height - maMargeB
    
    'On veut que unMaxYecran = unA * unVScroll.Min + unB
    'et unMin = unA * unVScroll.Max + unB d'où le calcul de unA et unB
    unA = (unMaxYecran - unMinYecran) / (unVScroll.Min - unVScroll.Max)
    unB = unMaxYecran - unA * unVScroll.Min
    
    'Calcul de la position en Y écran
    ConvertirVScrollValueEnYEcran = unA * uneVal + unB
End Function

Private Function ConvertirYEcranEnVScrollValue(unVScroll As VScrollBar, unY As Single) As Single
    'Fonction retournant la value des scroller vertical liés au début
    'et fin d'itinéraire chargé, en fonction de la coordonnée Y écran
    'du repère début ou fin
    Dim unMaxYecran As Single, unMinYecran As Single
    Dim unA As Single, unB As Single, uneVal As Single
    
    'Le max y écran correpondant au max y réel est en haut de fenêtre
    'c'est l'inverse pour le min y écran car Oy réel est orienté de bas en haut
    'alors que Oy écran est orienté de haut en bas
    unMaxYecran = maMargeB
    unMinYecran = ZoneDessin.Height - maMargeB
    
    'On veut que unMaxYecran = unA * unVScroll.Min + unB
    'et unMin = unA * unVScroll.Max + unB d'où le calcul de unA et unB
    unA = (unMaxYecran - unMinYecran) / (unVScroll.Min - unVScroll.Max)
    unB = unMaxYecran - unA * unVScroll.Min
    
    'Calcul de la position en Y écran
    uneVal = (unY - unB) / unA
    If uneVal < unVScroll.Min Then
        'Cas où on est en dessous du min du vscroll ===> on s'y met
        ConvertirYEcranEnVScrollValue = unVScroll.Min
    ElseIf uneVal > unVScroll.Max Then
        'Cas où on dépasse le max du vscroll ===> on s'y met
        ConvertirYEcranEnVScrollValue = unVScroll.Max
    Else
        'Autres cas uneVal entre le min et le max
        ConvertirYEcranEnVScrollValue = uneVal
    End If
End Function


Public Sub DeplacerTousRepIti()
    'Fonction déplacement tous les repères itinéraires, donc les ronds les symbolisant
    'et leur ligne de rappel respective et les labels début et fin
    'lorsque l'on clique sur l'ascenseur déplaçant le repère début
    Dim uneTransEcran As Single, uneTransReel As Single, unYreel As Single
    Dim uneVisu As Boolean
    Dim unYEcranRepDebAvant As Double, unNewYEcranRepDeb As Double
    Dim uneDistMaxReel As Single, uneDistMaxEcran As Single
    Dim unMinYecran As Single, unMaxYecran As Single
    
    'Indication d'une modif du vscroll repère début par programme (.value = qqchose)
    '==> DéplacerTousRepIti ne fera rien
    If monIsScrollByClick = False Then Exit Sub
    
    'Détermination des min/max écran
    'Variables servant pour la conversion abscisses curvilignes réelles en écran
    unMaxYecran = maMargeB
    unMinYecran = ZoneDessin.Height - maMargeB
    uneDistMaxReel = monMaxYReel - monMinYReel
    uneDistMaxEcran = ZoneDessin.Height - maMargeB * 2
    'Stockage du Y écran du repère début avant déplacement
    unYEcranRepDebAvant = LineRep(monIndRepItiDeb).Y1
    'Calcul de la nouvelle position du repère début
    unNewYEcranRepDeb = ConvertirVScrollValueEnYEcran(VScrollDebIti, VScrollDebIti.Value)
    LineRep(monIndRepItiDeb).Y1 = unNewYEcranRepDeb
    LabelDebIti.Top = unNewYEcranRepDeb - LabelDebIti.Height / 2
    'Calcul de la translation effectué
    uneTransEcran = unNewYEcranRepDeb - unYEcranRepDebAvant
    'Propagation à tous les repères itinéraires et aux controls les symbolisant
    LabelDebIti.Top = LabelDebIti.Top + uneTransEcran
    For i = 1 To monNbRepIti
        ShapeRep(i).Top = ShapeRep(i).Top + uneTransEcran
        If i = monIndRepItiDeb Then
            'On voit toujours le repère début
            'et son LineRep.Y1 a été modifié ci-dessus
            uneVisu = True
        Else
            LineRep(i).Y1 = LineRep(i).Y1 + uneTransEcran
            uneVisu = (LineRep(i).Y1 <= unMinYecran + EpsilonEcran And LineRep(i).Y1 >= unMaxYecran - EpsilonEcran)
        End If
        LineRep(i).Y2 = LineRep(i).Y1
        ShapeRep(i).Visible = uneVisu
        LineRep(i).Visible = uneVisu
        'Modification de l'abscisse curviligne réel stocké dans le tag
        'en convertissant l'abs curv écran en abs curv réelle
        unYreel = ConvertirEnReel(monMaxYReel, maMargeB - LineRep(i).Y1, uneDistMaxReel, uneDistMaxEcran)
        ShapeRep(i).Tag = Format(unYreel)
    Next i
    'Déplacement du label Fin du repère fin
    LabelFinIti.Top = LineRep(monIndRepItiFin).Y1 - LabelFinIti.Height / 2
    'Affichage des labels début/fin si la ligne de rappel du repère début/fin
    'est visible
    LabelDebIti.Visible = LineRep(monIndRepItiDeb).Visible
    LabelFinIti.Visible = LineRep(monIndRepItiFin).Visible
    'On grise le vscroll de modif du repère fin iti s'il n'est pas visible
    VScrollFinIti.Enabled = LineRep(monIndRepItiFin).Visible
    LabelFin.Enabled = LineRep(monIndRepItiFin).Visible
    'Déplacement des min et max réels de l'itinéraire d'une valeur de
    'translation réel en convertissant la translation écran
    uneTransReel = DonnerDistReel(uneTransEcran, uneDistMaxReel, uneDistMaxEcran)
    monMaxDReelZIti = monMaxDReelZIti - uneTransReel
    monMinDReelZIti = monMinDReelZIti - uneTransReel

    'Mise à jour de la position du scroller modifiant le repère iti réf de fin
    'Synchronisation de la nouvelle position dans le vscroll fin
    'Indication d'une modif des vscrollers repères début et fin par programme
    '(.value = qqchose)==> le change event sur vscroll ne fera rien
    monIsScrollByClick = False
    'Positionnement du vscroll pour le repère fin par rapport
    'au Y écran du repère fin d'itinéraire
    VScrollFinIti.Value = ConvertirYEcranEnVScrollValue(VScrollDebIti, LineRep(monIndRepItiFin).Y1)
    'Remise de l'Indication qu'une modif des vscrollers repères début et fin
    'par click souris sur leurs flèches est possible
    '==> le change event sur vscroll fera son travail dans ces cas
    monIsScrollByClick = True
End Sub

Public Sub EtirerTousRepIti()
    'Fonction étirant tous les repères itinéraires, donc les ronds les symbolisant
    'et leur ligne de rappel respective et les labels début et fin
    'lorsque l'on clique sur l'ascenseur déplaçant le repère fin
    'donc il peut y avoir contraction ou dilatation des abs curv ou y des repères
    Dim unCoefEtirageEcran As Double, unCoefEtirageReel As Single
    Dim unYreel As Single, unYEcranRepDeb As Double
    Dim unYEcranRepFinAvant As Double, unNewYEcranRepFin As Double
    Dim uneDistMaxReel As Single, uneDistMaxEcran As Single
    
    'Indication d'une modif du vscroll repère début par programme (.value = qqchose)
    '==> DéplacerTousRepIti ne fera rien
    If monIsScrollByClick = False Then Exit Sub
    
    'Stockage du Y écran du repère début
    unYEcranRepDeb = LineRep(monIndRepItiDeb).Y1
    'Détermination des min/max écran
    'Variables servant pour la conversion abscisses curvilignes réelles en écran
    uneDistMaxReel = monMaxYReel - monMinYReel
    uneDistMaxEcran = ZoneDessin.Height - maMargeB * 2
    'Stockage du Y écran du repère fin avant déplacement
    unYEcranRepFinAvant = LineRep(monIndRepItiFin).Y1
    'Calcul de la nouvelle position du repère fin
    unNewYEcranRepFin = ConvertirVScrollValueEnYEcran(VScrollFinIti, VScrollFinIti.Value)
    'LineRep(monIndRepItiFin).Y1 = unNewYEcranRepFin
    'Calcul de l'étirage effectué
    unCoefEtirageEcran = (unNewYEcranRepFin - unYEcranRepDeb) / (unYEcranRepFinAvant - unYEcranRepDeb)
    'Affichage du coefficient d'étirage = coef d'étalonnage
    TextCoefEta.Text = Format(CSng(TextCoefEta.Text) * unCoefEtirageEcran, "##0.0000")
    'Propagation à tous les repères itinéraires et aux controls les symbolisant
    'sauf le repère début d'itinéraire
    For i = 1 To monNbRepIti
        If i <> monIndRepItiDeb Then
            LineRep(i).Y1 = unYEcranRepDeb + (LineRep(i).Y1 - unYEcranRepDeb) * unCoefEtirageEcran
            LineRep(i).Y2 = LineRep(i).Y1
            ShapeRep(i).Top = LineRep(i).Y1 - ShapeRep(i).Height / 2
            'Modification de l'abscisse curviligne réel stocké dans le tag
            'en convertissant l'abs curv écran en abs curv réelle
            unYreel = ConvertirEnReel(monMaxYReel, maMargeB - LineRep(i).Y1, uneDistMaxReel, uneDistMaxEcran)
            ShapeRep(i).Tag = Format(unYreel)
        End If
    Next i
    'Mise au premier plan du repère fin
    ShapeRep(monIndRepItiFin).ZOrder 0
    'Déplacement du label fin du repère fin
    LabelFinIti.Top = unNewYEcranRepFin - LabelFinIti.Height / 2
    'Déplacement du max réel de l'itinéraire en convertissant la coordonnée écran
    'du repère fin en coordonnée relle
    monMaxDReelZIti = ConvertirEnReel(monMaxYReel, maMargeB - LineRep(monIndRepItiFin).Y1, uneDistMaxReel, uneDistMaxEcran)
End Sub

Private Sub ZoneDessin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Sélection des repères itinéraires ou des repères de parcours
    'Pour la sélection des repères itinéraires, on alimente la variable globale
    'privée monIndRepItiSel donnant l'indice du control ShapeRep cliqué,
    '0 si rien de sélectionner
    Dim unPar As Parcours, unD1 As Single, unX1 As Single, unY1 As Single
    Dim unX2 As Single, unY2 As Single, uneDist As Single, j As Long
    Dim unX As Single, unY As Single
    Dim uneDistMaxReelX As Single, uneDistMaxEcranX As Single
    Dim uneDistMaxReelY As Single, uneDistMaxEcranY As Single
    Dim unIndParChoisi As Integer
    
    'Si aucun itinéraire chargé, aucun sélection n'est possible
    'Le bouton changer début itinéraire apparait dés le premier chargement d'itinéraire
    If BtnChgDebIti.Enabled = False Then
        MsgBox "Les sélections graphiques des parcours sont possibles uniquement si un itinéraire est chargé.", vbInformation
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    
    'Détermination des min/max écran
    'Variables servant pour la conversion abscisses curvilignes réelles en écran
    uneDistMaxReelY = monMaxYReel - monMinYReel
    uneDistMaxEcranY = ZoneDessin.Height - maMargeB * 2
    uneDistMaxReelX = monMaxXReel - monMinXReel
    uneDistMaxEcranX = ZoneDessin.Width - monDec - maMargeD
    
    'Initialisation
    i = 0
    unIndParChoisi = 0
    'Vidage de la listbox des parcours trouvés
    frmChoixPar.Visible = False
    frmChoixPar.ListParTrouv.Clear
    TextNomPar.Text = "Aucun"
    TextDatePar.Text = ""
    
    If X >= ZoneDessin.Width - maMargeD + monDec And X <= ZoneDessin.Width Then
        'Cas où l'on clique à droite de l'axe Oy, il n'y a que des repères itinéraires
        'Recherche du rond , donc du repère itinéraire cliqué
        For i = 1 To monNbRepIti
            If X - ShapeRep(i).Left >= 0 And X - ShapeRep(i).Left <= ShapeRep(i).Width Then
                If Y - ShapeRep(i).Top >= 0 And Y - ShapeRep(i).Top <= ShapeRep(i).Height Then
                    'Cas où un rep iti est trouvé
                    If i = monIndRepItiDeb Or i = monIndRepItiFin Then
                        'Sélection du repère début ou fin d'iti impossible
                        MsgBox "La sélection de repère début ou fin d'itinéraire est impossible.", vbExclamation
                    Else
                        'Déselection de l'ancienne sélection, bordure n'est plus noire
                        ShapeRep(monIndRepItiSel).BorderColor = ShapeRep(monIndRepItiSel).BackColor
                        'Stockage de l'index de sélection et sélection,
                        'la bordure est mise en noir autour du rep iti trouvé
                        ShapeRep(i).BorderColor = QBColor(0)
                        monIndRepItiSel = i
                    End If
                    Exit For
                End If
            End If
        Next i
    Else
        'Autres cas, on est dans une partie où seul un parcours
        'issu du fichier MTB peut être cliqué
        i = 0 'pour indiquer de faire la déselection des rep iti
        
        'Conversion d'une distance de EpsilonEcran/2 twips en distance réelle sur OY et du Y écran
        unEpsilonY = DonnerDistReel(EpsilonEcran / 2, monMaxYReel - monMinYReel, uneDistMaxEcranY)
        unY = ConvertirEnReel(monMinYReel, ZoneDessin.Height - maMargeB - Y, uneDistMaxReelY, uneDistMaxEcranY)
        'Conversion d'une distance de EpsilonEcran/2 twips en distance réelle sur OY
        unEpsilonX = DonnerDistReel(EpsilonEcran / 2, monMaxXReel - monMinXReel, uneDistMaxEcranX)
        unX = ConvertirEnReel(monMinXReel, X - monDec, uneDistMaxReelX, uneDistMaxEcranX)
        'Epsilon pour la proximité en diagonale ou en projection
        unEpsilonXY = Sqr(unEpsilonX * unEpsilonX + unEpsilonY * unEpsilonY)
        
        'Recherche sur tous les parcours affichés du parcours cliqué
        For k = 1 To maColParcoursMTB.Count
            Set unPar = maColParcoursMTB(k)
            If unPar.monIsUtil Then
                unNbPoints = unPar.monNbPas
                'Conversion dixième de secondes en minutes
                unX1 = unPar.monFirstPas / 600
                uneDist = 0
                    
                For j = 2 To unNbPoints
                    'Recup des coordonnées x des points
                    'Conversion des dixièmes de seconde et des secondes en minutes
                    If j = unNbPoints Then
                        unX2 = unPar.monFirstPas / 600 + (unNbPoints - 2) * unPar.monPasMesure / 60 + unPar.monLastPas / 600
                    Else
                        unX2 = unPar.monFirstPas / 600 + (j - 1) * unPar.monPasMesure / 60
                    End If
                                    
                    'Calcul du Y
                    uneDist = uneDist + unPar.monTabDist(j - 1) * unPar.monCoefEta / 10
                    unY1 = uneDist 'Stockage pour l'incrémentation suivante
                    unY2 = uneDist + unPar.monTabDist(j) * unPar.monCoefEta / 10
                    
                    'Si point confondu on ne fait rien, on passe au suivant
                    unPtConfondu = (unX1 = unX2 And unY1 = unY2)
                    If unPtConfondu = False Then
                        'Recherche si on a cliqué prés du segment M1(x1,y1)-M2(x2,y2)
                        'D'abord on regarde si X1 < X < X2, puis Y1 < Y < Y2 et enfin
                        'si la distance à la droite passant par M1 et M2 est < espilon
                        If (unX1 - unEpsilonX < unX) And (unX < unX2 + unEpsilonX) Then
                            If (unY1 - unEpsilonY < unY) And (unY < unY2 + unEpsilonY) Then
                                'Calcul de la distance à la droite pasant par M1M2
                                'ax+by+c = 0, a = y2-y1, b=x1-x2, c = x2y1-x1y2
                                uneDistM1M2 = (unY2 - unY1) * unX + (unX1 - unX2) * unY + (unX2 * unY1 - unX1 * unY2)
                                If Abs(uneDistM1M2) / Sqr((unX2 - unX1) * (unX2 - unX1) + (unY2 - unY1) * (unY2 - unY1)) < unEpsilonXY Then
                                    'Cas où le parcours est clqué
                                    'Stockage du parcours et des x et y écran du repère sélectionné
                                    unIndParChoisi = k
                                    'Ajout dans la listbox des parcours trouvés
                                    frmChoixPar.ListParTrouv.AddItem unPar.monNom + " (" + Mid(unPar.monJourSemaine, 1, 2) + " " + Format(unPar.maDate) + " " + Mid(Format(unPar.monHeureDebut), 1, 5) + ")"
                                    frmChoixPar.ListParTrouv.ItemData(frmChoixPar.ListParTrouv.NewIndex) = k
                                    'Sortie du for bouclant sur les points du parcours
                                    'on passe au parcours suivant
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                    'Pour l'incrémentation suivante
                    unX1 = unX2
                Next j
            End If
        Next k
        If unIndParChoisi = 0 Then
            MsgBox "Aucun parcours n'a été sélectionné.", vbInformation
        End If
    End If
    
    'Si aucun sélection de repère, on déselectionne le dernier rep iti sélectionné
    If i = 0 Or i = monNbRepIti + 1 Then
        ShapeRep(monIndRepItiSel).BorderColor = ShapeRep(monIndRepItiSel).BackColor
        monIndRepItiSel = 0
    End If
    
    'Si plus d'un parcours sélectionné, on affichage leur nom dans une fenêtre
    'de choix qui apparait au centre de l'écran
    If frmChoixPar.ListParTrouv.ListCount > 1 Then
        CentrerFenetreEcran frmChoixPar
        frmChoixPar.Show vbModal
        If Tag = "" Then
            'Cas où aucun parcours choisi
            '===> click sur bouton Annuler de la fenêtre choix parcours
            unIndParChoisi = 0
        Else
            'Cas où un parcours  a été choisi
            unIndParChoisi = CInt(Tag)
        End If
    End If
            
    If unIndParChoisi > 0 Then
        'Cas où un parcours a été choisi par click sur un repère
        'ou par choix parmi plusieurs et on l'affiche en trait épais
        TextNomPar.Text = maColParcoursMTB(unIndParChoisi).monNom
        TextDatePar.Text = maColParcoursMTB(unIndParChoisi).monJourSemaine + " " + Format(maColParcoursMTB(unIndParChoisi).maDate) + " " + Format(maColParcoursMTB(unIndParChoisi).monHeureDebut)
        If frmChoixPar.ListParTrouv.ListCount = 1 Then
            'On redessine le parcours choisi en trait gros uniquement
            'si un seul parcours trouvé, sinon on redessine plusieurs fois
            MontrerParcoursChoisi unIndParChoisi
        End If
        'Stockage de l'indice du parcours sélection
        monIndAncienParChoisi = unIndParChoisi
    ElseIf monIndAncienParChoisi > 0 Then
        'Cas où rien de sélectionner par click, mais où un parcours
        'avait été cliqué avant, on met l'ancienne sélection à 0 pour
        'tout redessiner en trait fin
        monIndAncienParChoisi = 0
        MontrerParcoursChoisi monIndAncienParChoisi
    ElseIf Mid(monBtnClick, 1, 7) = "Annuler" And Len(monBtnClick) > 7 Then
        'Longueur > 7 car monBtnclcik peut ne valoir qu'annuler si on a utilisé
        'd'autres fenêtres avant qui avait des boutons ok et annuler remplissant
        'la variable globale monBtnClick, frmChoixPar donne Annuler plus l'index du
        'parcours sélectionné ou -1 si aucun parcours n'est sélectionné
        If CInt(Mid(monBtnClick, 8)) > -1 Then
            'On redessine tous les parcours choisi en trait fin en disant qu'aucun n'a
            'été trouvé
            monIndAncienParChoisi = 0
            MontrerParcoursChoisi monIndAncienParChoisi
            monBtnClick = "Annuler" + Format(-1)
        End If
    End If
    
    MousePointer = vbDefault
End Sub

Public Sub MontrerParcoursChoisi(unIndParChoisi As Integer)
    'Dessin des axes
    DessinerAxes
    'Dessin des parcours en mettant en trait épais le parcours choisi
    DessinerDesParcours ZoneDessin, maColParcoursMTB, maMargeD, monDec, maMargeB, maMargeB, monMinXReel, monMaxXReel, monMinYReel, monMaxYReel, unIndParChoisi
End Sub
