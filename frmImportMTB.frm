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
      Caption         =   "Modifier l'itin�raire"
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
         Caption         =   "Coefficient d' �tirement : "
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
         Caption         =   "D�but"
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
      Caption         =   "L�gende"
      Height          =   780
      Left            =   8760
      TabIndex        =   13
      Top             =   5760
      Width           =   2055
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rep�re parcours"
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
         Caption         =   "Rep�re itin�raire"
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
      Caption         =   "Changer d�but d'itin�raire"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   2490
      Width           =   2055
   End
   Begin VB.CommandButton BtnChgFinIti 
      Caption         =   "Changer fin d'itin�raire"
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
      Caption         =   "Zoom cadr� sur l'itin�raire"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8760
      TabIndex        =   4
      Top             =   1530
      Width           =   2055
   End
   Begin VB.CommandButton BtnImport 
      Caption         =   "Importer le parcours cadr�"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   1050
      Width           =   2055
   End
   Begin VB.CommandButton BtnCharger 
      Caption         =   "Charger l'itin�raire..."
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
         Caption         =   "D�but"
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
      Caption         =   "Couleur du Rep�re d�but itin�raire"
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
      Caption         =   "Couleur du Rep�re fin itin�raire"
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
      Caption         =   "Parcours � importer : "
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   540
      Width           =   1500
   End
   Begin VB.Label LabelItiRef 
      AutoSize        =   -1  'True
      Caption         =   "Itin�raire : "
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
'Variable indiquant si les rep�res d�but et fin de l'itin�raire charg�e
'ont �t� modifi�
Private maModifRepItiDeb As Boolean
Private maModifRepItiFin As Boolean
'Collection stockant les indices des parcours qui auront �t� import�s
Private maColParImport As New Collection
'Variables stockant les min et max des temps et vitesse
'de l'itin�raire de r�f�rence charg�
Private monMaxT As Single, monMaxV As Single
Private monMinT As Single, monMinV As Single
'Variables stockant les max et min r�els en Y suivant le zoom tout ou iti
Private monMaxYReel As Single, monMinYReel As Single
'Variables stockant les max et min r�els en X suivant le zoom tout ou iti
Private monMaxXReel As Single, monMinXReel As Single
'Variables stockant les min et max r�els en temps et distances
'en zoom tout et en zoom itin�raire
Private monMaxTReelZTout As Single, monMaxDReelZTout As Single
Private monMinTReelZTout As Single, monMinDReelZTout As Single
Private monMaxTReelZIti As Single, monMaxDReelZIti As Single
Private monMinTReelZIti As Single, monMinDReelZIti As Single
'Variable indiquant si les scrollers verticaux des rep�res d�but et fin ont
'�t� modifi�s par programme (.value = qqchose) ou par click sur leurs fl�ches
Private monIsScrollByClick As Boolean


Private Sub BtnAide_Click()
    's'il n'y pas de fichier d'aide pour le projet, afficher un message � l'utilisateur
    'vous pouvez d�finir le fichier d'aide de votre application dans la bo�te
    'de dialogue de propri�t�s du projet
    If Len(App.HelpFile) = 0 Then
        MsgBox "Impossible d'afficher le sommaire de l'aide. Il n'y a pas d'aide associ�e � ce projet.", vbInformation, Me.Caption
    Else
        'Lance l'aide du bon contexte
        frmMain.dlgCommonDialog.HelpCommand = cdlHelpContext
        frmMain.dlgCommonDialog.HelpContext = HelpContextID
        frmMain.dlgCommonDialog.ShowHelp  ' affiche la rubrique
    End If
End Sub

Private Sub BtnCharger_Click()
    'Ouvre l'itin�raire contenue dans le fichier choisi dans cette fonction
    'et affiche ses rep�res
    Dim uneAbsCurv As Long, unTypeIco As Byte
    Dim unNbCarLig1a5 As Long
    Dim uneString As String, uneString1 As String
    Dim unTabSng(1 To NbPasMax) As Single 'Pour stocker les lectures de single
    Dim unCheckSection As Integer, unIndRepDeb As Integer, unIndRepFin As Integer
    Dim unNomFich As String

    'Si protection invalide on ne fait rien
    'If ProtectCheck(2) <> 0 Then Exit Sub
    
    'Vidage de la collection des indices de parcours import�s
    ViderCollection maColParImport
    
    'Fermeture du fichier charg� �ventuel
    If monFichId > 0 Then Close #monFichId
    
    'Choix du fichier itin�raire
    unNomFich = frmMain.ChoisirFichier(MsgOpen, MsgMitFile, CurDir)
    If unNomFich = "" Then Exit Sub 'Si aucun fichier choisi ou annuler, on sort
    
    'Affichage du sablier pointeur souris d'attente
    Me.MousePointer = vbHourglass
    
    'Initialisation
    monMaxDReelZIti = -100
    monMinDReelZIti = 10000000
    
    'Afichage du nom du fichier itin�raire
    TextFichItiRef.Text = unNomFich
    'Affichage de la fin du nom de fichier si d�passe zone texte
    'en pla�ant le curseur en fin de texte
    TextFichItiRef.SelStart = Len(TextFichItiRef.Text)
        
    If ShapeRep.Count > 1 And BtnZoomTout.Enabled = True Then
        'Cas o� on charge un autre fichier itin�raire � la place de celui
        'd�j� charg�, donc des shaperep existe d�j� sauf au premier chargement d'un
        'fichier itin�raire ==> ReDessin des parcours issus du fichier MTB = campagne
        'de mesures en zoom tout les parcours si on �tait en zoom cadr� sur itin�raire
        ZoneDessin.Cls
        DessinerAxes
        DessinerDesParcours ZoneDessin, maColParcoursMTB, maMargeD, monDec, maMargeB, maMargeB, monMinTReelZTout, monMaxTReelZTout, monMinDReelZTout, monMaxDReelZTout, monIndAncienParChoisi
    End If
    
    'Lecture du fichier .mit
    ' Active la routine de gestion d'erreur.
    'MsgBox "Suppression du On Error GoTo ErreurLecture"
    On Error GoTo ErreurReadIti
    
    'Ouverture du fichier en lecture lock�e pour �viter deux ouvertures
    monFichId = FreeFile(0) 'renvoi d'un nombre entre 1 et 255
    Open unNomFich For Input Lock Read Write As #monFichId
        
    'Lecture de l'ent�te des fichiers *.mit
    Input #monFichId, uneString
    If uneString <> "Fichier " + App.Title Then
        'Cas d'un fichier qui n'est pas un fichier MiTemps
        '===> Fermeture du fichier.
        Close #monFichId
        MsgBox MsgErreur + MsgFileNotFile + App.Title + "version " + App.Major + "." + App.Minor, vbCritical
    Else
        'Cas d'un fichier MiTemps *.Mit de la version 3.0
        '1�re ligne du fichier MIT = "Fichier MiTemps"
        
        'Lecture des libell�s des conditions m�t�o
        Input #monFichId, uneString
        
        'Lecture des min et max total en distance (parcours complet sans section)
        Input #monFichId, monMinDReelZIti, monMaxDReelZIti
        'R�cup�ration des min et max en distance, vitesse et temps
        Input #monFichId, unTabSng(1), unTabSng(2), monMinV, monMaxV, monMinT, monMaxT
        
        'Lecture du nom de l'itin�raire et de sa longueur
        Input #monFichId, uneString, uneString1
        
        'R�cup�ration des donn�es de la section de travail
        Input #monFichId, unCheckSection, unIndRepDeb, unIndRepFin
        
        'Stockage de la position de lecture juste avant la lecture du nombre
        'de rep�res et parcours, on s'en sert lors de l'importation pour �crire
        'dans le fichier mit de l'itin�raire charg� le nombre de parcours + 1
        unNbCarLig1a5 = Seek(monFichId)
        
        'R�cup�ration du nombre de rep�res et de parcours
        Input #monFichId, monNbRepIti, monNbPar
        'Sauvegarde de l'ancien nombre de rep iti, avant le d�coupage
        'On s'en sert dans la fonction CouperParcoursEntreD1D2 plus bas
        'pour avoir le bon nombre de rep�res en Control Shape et �viter des
        'plantages dans la fonction CouperParcoursEntreD1D2
        monSaveNbRepIti = monNbRepIti
        'R�cup�ration et cr�ation des rep�res
        uneAbsMax = -100
        uneAbsMin = 10000000
        For i = 1 To monNbRepIti
            Input #monFichId, uneString, uneString1, uneAbsCurv, unTypeIco
            If i > monNbRepTot Then
                'Cas o� il faut cr�er un contr�le rep de plus
                'sinon on modifie juste ses param�tres
                Load ShapeRep(i)
                Load LineRep(i)
                monNbRepTot = monNbRepTot + 1
            End If
            ShapeRep(i).Visible = True
            ShapeRep(i).BackColor = ShapeRepIti.BackColor 'Couleur du rep interm�diaire
            ShapeRep(i).BorderColor = ShapeRepIti.BorderColor
            ShapeRep(i).Tag = Format(uneAbsCurv)
            LineRep(i).Visible = True
            'Recherche de l'abs curv min et max pour les rep�res
            'd�but et fin d'itin�raire
            If uneAbsCurv >= uneAbsMax Then
                uneAbsMax = uneAbsCurv
                monIndRepItiFin = i
            End If
            If uneAbsCurv <= uneAbsMin Then
                uneAbsMin = uneAbsCurv
                monIndRepItiDeb = i
            End If
        Next i
        
        'Dessin des rep�res itin�raire au zoom tout
        DessinerRepereIti ZoneDessin, maMargeD, monDec, maMargeB, maMargeB, monMinTReelZTout, monMaxTReelZTout, monMinDReelZTout, monMaxDReelZTout
        'Indication du rep�re d�but et fin de l'itin�raire charg�
        ShapeRep(monIndRepItiDeb).BackColor = ShapeRepDeb.BackColor
        ShapeRep(monIndRepItiDeb).BorderColor = ShapeRepDeb.BorderColor
        ShapeRep(monIndRepItiFin).BackColor = ShapeRepFin.BackColor
        ShapeRep(monIndRepItiFin).BorderColor = ShapeRepFin.BorderColor
    End If
    
    'Masquage des autres contr�les rep�res qui ne servent plus
    For i = monNbRepIti + 1 To monNbRepTot
        ShapeRep(i).Visible = False
        LineRep(i).Visible = False
    Next i
    
    'Activation du bouton Zoom iti et des boutons de changment et de fin
    'de l'itin�raire et du parcours et d�activation du bouton zoom tout les parcours
    'car c'est le zoom actuel
    BtnZoomTout.Enabled = False
    BtnZoomIti.Enabled = True
    BtnChgDebIti.Enabled = True
    BtnChgFinIti.Enabled = True
    BtnImport.Enabled = True
    
    'Activation des scrollers permettant la modif des d�buts et fin d'itin�raire
    FrameDebFinIti.Enabled = True
    LabelDeb.Enabled = True
    LabelFin.Enabled = True
    VScrollDebIti.Enabled = True
    VScrollFinIti.Enabled = True
    
    'Stockage des min et max r�els en X et Y
    monMaxYReel = monMaxDReelZTout
    monMinYReel = monMinDReelZTout
    monMaxXReel = monMaxTReelZTout
    monMinXReel = monMinTReelZTout
    
    'Initialisation de l'index du rep�re iti s�lectionn�, 0 = rien de s�lectionner
    monIndRepItiSel = 0
    
    'Initialisation du coef d'�talonnage (4 chiffres apr�s la virgule comme dans
    'les fichiers MTB MiTemps) et on d�grise son affichage
    TextCoefEta.Text = Format(1, "###.0000")
    LabelCoefEta.Enabled = True
    TextCoefEta.Enabled = True
    
    'Stockage de la position de lecture juste avant la lecture du nombre
    'de rep�res et parcours, on s'en sert lors de l'importation pour �crire
    'dans le fichier mit de l'itin�raire charg� le nombre de parcours + 1
    monNbCarLig1a5 = unNbCarLig1a5
        
    'Initialisation de l'�tat d'une modif du rep�re d�but ou/et fin
    maModifRepItiDeb = False
    maModifRepItiFin = False
    'Affichage du pointeur souris par d�faut
    Me.MousePointer = vbDefault
    ' Quitte pour �viter le gestionnaire d'erreur et on le d�sactive.
    On Error GoTo 0
    Exit Sub
    
    ' Routine de gestion d'erreur qui �value le num�ro d'erreur.
ErreurReadIti:
    
    ' Traite les autres situations ici...
    unMsg = MsgOpenError + unNomFich + Chr(13) + Chr(13) + MsgErreur + Format(Err.Number) + " : " + Err.Description
    If Err.Number = 70 Then unMsg = unMsg + " (" + UCase(MsgDejaOpen) + ")"
    MsgBox unMsg, vbCritical
    'fermeture du fichier
    Close #monFichId
    Me.MousePointer = vbDefault
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub BtnChgDebIti_Click()
    'Changement de localistation du rep�re d�but d'itin�raire
    If monIndRepItiSel = 0 Then
        'aucun rep iti s�lectionn�
        MsgBox "Il faut d'abord choisir un rep�re de l'itin�raire en cliquant dessus.", vbInformation
    ElseIf CSng(ShapeRep(monIndRepItiSel).Tag) >= CSng(ShapeRep(monIndRepItiFin).Tag) Then
        'Cas o� Abs curv rep s�lectionn� pour �tre rep�re d�but est
        ' > abs curv rep�re fin iti ==> Impossible
        MsgBox "Le rep�re d�but d'itin�raire ne peut pas �tre apr�s le rep�re fin d'itin�raire.", vbExclamation
    Else
        'Cas o� l'on peut modifier le rep�re d�but
        'Affichage du nouveau rep�re d�but avec sa couleur et son label
        ShapeRep(monIndRepItiSel).BackColor = ShapeRep(monIndRepItiDeb).BackColor
        ShapeRep(monIndRepItiSel).BorderColor = ShapeRep(monIndRepItiDeb).BorderColor
        LabelDebIti.Top = LineRep(monIndRepItiSel).Y1 - LabelDebIti.Height / 2
        'Mise en couleur normale de l'ancien rep�re d�but
        ShapeRep(monIndRepItiDeb).BackColor = ShapeRepIti.BackColor
        ShapeRep(monIndRepItiDeb).BorderColor = ShapeRepIti.BorderColor
        'Stockage du nouveau index de rep�re d�but
        monIndRepItiDeb = monIndRepItiSel
        'Mise � jour du min pour le zoom cadr� sur l'itin�raire
        monMinDReelZIti = CSng(ShapeRep(monIndRepItiSel).Tag)
        
        'Synchronisation de la nouvelle position dans le vscroll d�but
        'Indication d'une modif des vscrollers rep�res d�but et fin par programme
        '(.value = qqchose)==> le change event sur vscroll ne fera rien
        monIsScrollByClick = False
        'Positionnement du vscroll pour le rep�re d�but par rapport
        'au Y �cran du rep�re d�but d'itin�raire
        VScrollDebIti.Value = ConvertirYEcranEnVScrollValue(VScrollDebIti, LineRep(monIndRepItiSel).Y1)
        'Remise de l'Indication qu'une modif des vscrollers rep�res d�but et fin
        'par click souris sur leurs fl�ches est possible
        '==> le change event sur vscroll fera son travail dans ces cas
        monIsScrollByClick = True
        
        'Indication d'une modif du rep�re d�but
        maModifRepItiDeb = True
    End If
End Sub

Private Sub BtnChgFinIti_Click()
    'Changement de localistation du rep�re fin d'itin�raire
    If monIndRepItiSel = 0 Then
        'aucun rep iti s�lectionn�
        MsgBox "Il faut d'abord choisir un rep�re de l'itin�raire en cliquant dessus.", vbInformation
    ElseIf CSng(ShapeRep(monIndRepItiSel).Tag) <= CSng(ShapeRep(monIndRepItiDeb).Tag) Then
        'Cas o� Abs curv rep s�lectionn� pour �tre rep�re fin est
        ' < abs curv rep�re d�but iti ==> Impossible
        MsgBox "Le rep�re fin d'itin�raire ne peut pas �tre avant le rep�re d�but d'itin�raire.", vbExclamation
    Else
        'Cas o� l'on peut modifier le rep�re fin
        'Affichage du nouveau rep�re fin avec sa couleur et son label
        ShapeRep(monIndRepItiSel).BackColor = ShapeRep(monIndRepItiFin).BackColor
        ShapeRep(monIndRepItiSel).BorderColor = ShapeRep(monIndRepItiFin).BorderColor
        LabelFinIti.Top = LineRep(monIndRepItiSel).Y1 - LabelFinIti.Height / 2
        'Mise en couleur normale de l'ancien rep�re fin
        ShapeRep(monIndRepItiFin).BackColor = ShapeRepIti.BackColor
        ShapeRep(monIndRepItiFin).BorderColor = ShapeRepIti.BorderColor
        'Stockage du nouveau index de rep�re fin
        monIndRepItiFin = monIndRepItiSel
        'Mise � jour du max pour le zoom cadr� sur l'itin�raire
        monMaxDReelZIti = CSng(ShapeRep(monIndRepItiSel).Tag)
        
        'Synchronisation de la nouvelle position dans le vscroll fin
        'Indication d'une modif des vscrollers rep�res d�but et fin par programme
        '(.value = qqchose)==> le change event sur vscroll ne fera rien
        monIsScrollByClick = False
        'Positionnement du vscroll pour le rep�re fin par rapport
        'au Y �cran du rep�re fin d'itin�raire
        VScrollFinIti.Value = ConvertirYEcranEnVScrollValue(VScrollDebIti, LineRep(monIndRepItiSel).Y1)
        'Remise de l'Indication qu'une modif des vscrollers rep�res d�but et fin
        'par click souris sur leurs fl�ches est possible
        '==> le change event sur vscroll fera son travail dans ces cas
        monIsScrollByClick = True
        
        'Indication d'une modif du rep�re fin
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
        'aucun parcours s�lectionn�
        MsgBox "Il faut d'abord choisir un parcours en cliquant sur sa courbe Distance / Temps.", vbInformation
    Else
        'R�cup�ration du parcours � importer et des abscisses curvilignes
        'des rep�res d�but et fin de l'itin�raire charg�
        Set unParToImport = maColParcoursMTB(monIndAncienParChoisi)
        'R�cup�ration des abs curv de d�but et de fin
        uneD1 = CSng(Format(ShapeRep(monIndRepItiDeb).Tag))
        uneD2 = CSng(Format(ShapeRep(monIndRepItiFin).Tag))
        'On fait en sorte que D1 <= D2
        If uneD1 > uneD2 Then
            uneD0 = uneD1
            uneD1 = uneD2
            uneD2 = uneD0
        End If
        'Test si les rep�res d�but ou/et fin de l'itin�raire ont �t� chang�
        'On invite l'utilisateur � choisir un nouveau fichier d'itin�raire qui ne
        'contiendra aucun parcours et uniquement les rep�res entre les nouveaux
        'd�but et fin et qui servira � stocker les parcours import�s par coupure
        'entre les nouveaux rep�res d�but et fin
        If maModifRepItiDeb Or maModifRepItiFin Then
            'Message d'avertissement et de confirmation de continuation d'action
            unMsg = "Les rep�res d�but et/ou fin de l'itin�raire de r�f�rence ont �t� chang�s." + Chr(13) + Chr(13)
            unMsg = unMsg + "Pour �viter d'avoir des parcours avec un nombre de rep�res diff�rents," + Chr(13)
            unMsg = unMsg + "vous allez devoir choisir un nouveau fichier itin�raire qui ne contiendra" + Chr(13)
            unMsg = unMsg + "aucun parcours et uniquement les rep�res entre les nouveaux rep�res" + Chr(13)
            unMsg = unMsg + "d�but et fin et qui servira � stocker les parcours qui seront import�s" + Chr(13)
            unMsg = unMsg + "par coupure entre ces nouveaux rep�res d�but et fin."
            unMsg = unMsg + Chr(13) + Chr(13) + "Voulez-vous continuer ?"
            If MsgBox(unMsg, vbYesNo + vbQuestion) = vbNo Then Exit Sub
            
            'Stockage de l'ancien fichier itin�raire charg�
            unNomFich0 = TextFichItiRef.Text
            'Demande du nouveau fichier MIT de stockage
            unNomFich = frmMain.ChoisirFichier(MsgSaveAs, MsgMitFile, CurDir)
            If unNomFich = "" Then Exit Sub 'Si aucun fichier choisi ou annuler, on sort
            
            'Fermeture et r�ouverture en mode verrouill� pour
            '�viter deux ouvertures de l'ancien fichier itin�raire
            'ainsi sa lecture repart de la 1�re ligne
            unOldFichId = FreeFile(0)
            Close #monFichId
            Open unNomFich0 For Input Lock Read Write As #unOldFichId
            'Lecture des lignes jusqu'� la fin des rep�res et stockage des deux
            'premi�res lignes pour insertion dans le new MIT
            Input #unOldFichId, unTextLine1
            Input #unOldFichId, unTextLine2
            Input #unOldFichId, unR1, unR2
            Input #unOldFichId, unR1, unR2, unR3, unR4, unR5, unR6
            'R�cup�ration du nom et la longueur de l'itin�raire
            Input #unOldFichId, unNomIti, uneLongIti
            Input #unOldFichId, unI1, unI2, unI3
            Input #unOldFichId, unI1, unI2
            'R�cup�ration des rep�res se trouvant entre les rep�res d�but et fin et
            'du nombre de parcours que l'on met � 0 et calcul de ce nombres de rep�res
            unNbPar = 0
            unNbRepIti = 0 'Initialisation pour le calcul du nombre de rep�res
            For j = 1 To monNbRepIti
                Input #unOldFichId, unNomLong, unNomCourt, uneAbsCurv, unTypeIco
                If uneAbsCurv >= uneD1 - mesOptions.monEcartMax And uneAbsCurv <= uneD2 + mesOptions.monEcartMax Then
                    'On incr�mente le nombre de rep�res total
                    unNbRepIti = unNbRepIti + 1
                    'Calcul des nouveaux indices des rep�res d�but et fin
                    If Abs(uneAbsCurv - uneD1) < mesOptions.monEcartMax Then
                        unIndRepItiDeb = unNbRepIti
                    ElseIf Abs(uneAbsCurv - uneD2) < mesOptions.monEcartMax Then
                        unIndRepItiFin = unNbRepIti
                    End If
                    'On ajoute en fin de liste � chaque fois
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
            'Fermeture du fichier itin�raire pr�c�demment charg�
            Close #unOldFichId
            'Ouverture en mode output du nouveau fichier itin�raire pour �crire dedans
            Open unNomFich For Output As #monFichId
            'Ecriture des deux premi�res lignes du mit de d�part
            Write #monFichId, unTextLine1
            Write #monFichId, unTextLine2
            'Ecriture du min et du max total en distance
            Write #monFichId, 0, CLng(uneD2 - uneD1)
            'Calcul des min et max en distance, vitesse et dur�e
            'en convertissant D1 et D2 des m�tres en d�cim�tres
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
            'Conversion du temps des dixi�mes de secondes en minutes
            uneDuree = uneDuree / 600
            If unMaxT < uneDuree Then unMaxT = uneDuree
            'Ecriture des min et max en distance, vitesse et dur�e
            Write #monFichId, unMinD, CLng(unMaxD), unMinV, CLng(unMaxV), unMinT, CLng(unMaxT * 10) / 10
            'Ecriture du nom et de la longueur du nouvel itin�raire de r�f�rence
            'Write #monFichId, unNomIti, Format(unMaxD)
            uneStrTmp = unNomIti + " import� entre " + Format(uneD1) + " et " + Format(uneD2) + " m�tres"
            Write #monFichId, uneStrTmp, Format(unMaxD)
            'Ecriture de la section de travail, pas de section mais d�but et fin
            'correspondant aux rep�res d�but et fin du nouvel itin�raire donc
            'aux nouveaux indices calcul�s ci-dessus car les rep�res ne sont pas
            'class�es par ordre croissant
            uneSectionDefinie = 0
            Write #monFichId, uneSectionDefinie, unIndRepItiDeb, unIndRepItiFin
            'Stockage de la position de lecture juste avant la lecture du nombre
            'de rep�res et parcours, on s'en sert lors de l'importation pour �crire
            'dans le fichier mit de l'itin�raire charg� le nombre de parcours + 1
            monNbCarLig1a5 = Seek(monFichId)
            'Ecriture du nb de rep�res et de parcours et des rep�res qui ont �t�
            'trouv�s entre les rep�res d�but et fin
            Print #monFichId, Format(unNbRepIti) + "," + Format(unNbPar, "000#")
            For j = 0 To unNbRepIti - 1
                Write #monFichId, uneColString(4 * j + 1), uneColString(4 * j + 2), CLng(uneColString(4 * j + 3)), CInt(uneColString(4 * j + 4))
            Next j
            'On supprime la collection
            ViderCollection uneColString
            Set uneColString = Nothing
                        
            'Fermeture et r�ouverture en mode verrouill� pour
            '�viter deux ouvertures du nouveau fichier itin�raire
            Close #monFichId
            Open unNomFich For Input Lock Read Write As #monFichId
            'Affichage en haut du nouveau fichier d'itin�raire de r�f�rence
            TextFichItiRef.Text = unNomFich
            'Initialisation de l'�tat d'une modif du rep�re d�but ou/et fin
            maModifRepItiDeb = False
            maModifRepItiFin = False
            'Vidage de la collection des parcours d�j� import�
            ViderCollection maColParImport
            'Mise � jour du nouveau nombre de parcours et de rep�res
            monNbRepIti = unNbRepIti
            monNbPar = unNbPar
        End If
        
        'Test de la pr�sence de l'indice du parcours choisi dans la
        'collection des indices des parcours import�s
        '0n pr�vient de l'import mulitple du m�me parcours pendant la session
        'de travail, mais cela peut servir si on d�coupe par rapport � des lignes
        'de bus, d'o� on laisse l'utilisateur choisir
        unParDejaImport = vbYes
        For i = 1 To maColParImport.Count
            If monIndAncienParChoisi = maColParImport(i) Then
                unMsg = "Le parcours " + maColParcoursMTB(monIndAncienParChoisi).monNom + " mesur� le "
                unMsg = unMsg + maColParcoursMTB(monIndAncienParChoisi).monJourSemaine + " " + Format(maColParcoursMTB(monIndAncienParChoisi).maDate) + " � " + Format(maColParcoursMTB(monIndAncienParChoisi).monHeureDebut)
                unMsg = unMsg + " a d�j� �t� import� dans le fichier " + TextFichItiRef.Text
                unMsg = unMsg + Chr(13) + Chr(13) + "Voulez-vous continuer ?"
                unParDejaImport = MsgBox(unMsg, vbYesNo + vbQuestion)
                'unParDejaImport = True
                Exit For
            End If
        Next i
        'If unParDejaImport = False Then
        If unParDejaImport = vbYes Then
            unMsg = "Voulez-vous importer le parcours " + maColParcoursMTB(monIndAncienParChoisi).monNom + " mesur� le "
            unMsg = unMsg + maColParcoursMTB(monIndAncienParChoisi).monJourSemaine + " " + Format(maColParcoursMTB(monIndAncienParChoisi).maDate) + " � " + Format(maColParcoursMTB(monIndAncienParChoisi).monHeureDebut)
            unMsg = unMsg + " en le nommant " + Chr(34) + TextNomPar.Text + Chr(34) + " dans le fichier " + TextFichItiRef.Text + " ?"
            If MsgBox(unMsg, vbQuestion + vbYesNo) = vbNo Then
                'Sortie directe si on ne veut pas importer le parcours s�lectionn�
                Exit Sub
            End If
                                    
            'Coupure du parcours � importer entre les rep�res d�but et fin
            'de l'itin�raire charg� et on met les donn�es dans le parcours
            'donn�e par la variable globale monParToImport
            unCoefEtirement = Format(CSng(TextCoefEta.Text))
            If CouperParcoursEntreD1D2(unParToImport, Me, monSaveNbRepIti, monIndRepItiDeb, monIndRepItiFin, unCoefEtirement) = False Then
                'Si la coupure ne peut se faire, on sort sans rien faire
                Exit Sub
            End If
            
            ' Active la routine de gestion d'erreur.
            'MsgBox "Suppression du On Error GoTo ErreurRWIti"
            On Error GoTo ErreurRWIti
            
            'Fermeture du fichier itin�raire charg� pour r�ouverture en append apr�s
            Close #monFichId
            
            'Ouverture du fichier itin�raire charg� en mode ajout pour ajouter les
            'donn�es du parcours import�
            unNomFich = TextFichItiRef.Text
            Open unNomFich For Append As #monFichId
            monParToImport.monNom = TextNomPar.Text
            EcrireDonneesParcoursDansFichierMIT monFichId, monParToImport
            
            'On se met sur la ligne o� se trouve le nombre de rep�res et de parcours
            'pour ajouter +1 au nombre de parcours en le formattant sur 4 caract�res
            'pour �crire juste dans sa place (4 caract�res lors de la sauvegarde du
            'fichier MIT).
            'Cette position est donn�e par monNbCarLig1a5 d�finie dans la fonction
            'BtnCharger_Click liant le fichier itin�raire charg� ou dans cette
            'proc�dure plus haut dans la partie changeant le fichier en cas de modif
            'de rep�re d�but et/ou fin
            Seek #monFichId, monNbCarLig1a5
            unTextLine = Format(monNbRepIti) + "," + Format(monNbPar + 1, "000#")
            Print #monFichId, unTextLine
            
            'Fermeture et r�ouverture en mode verrouill� pour
            '�viter deux ouvertures
            Close #monFichId
            Open unNomFich For Input Lock Read Write As #monFichId
            'Ajout dans la collection des indices des parcours import�s
            maColParImport.Add monIndAncienParChoisi
            'Incr�mentation du nombre de parcours total car ici l'importation
            'est finie et c'est bien pass�e, donc on peut faire +1
            monNbPar = monNbPar + 1
            MsgBox "Importation r�ussie", vbInformation
        End If
    End If
    
    'Sortie pour �viter la gestion d'erreur
    On Error GoTo 0
    Exit Sub
    
    ' Routine de gestion d'erreur qui �value le num�ro d'erreur.
ErreurRWIti:
    
    ' Traite les autres situations ici...
    unMsg = MsgOpenError + unNomFich + Chr(13) + Chr(13) + MsgErreur + Format(Err.Number) + " : " + Err.Description
    If Err.Number = 70 Then unMsg = unMsg + " (" + UCase(MsgDejaOpen) + ")"
    MsgBox unMsg, vbCritical
    'Fermeture et r�ouverture en mode verrouill�
    Close #monFichId
    Open unNomFich For Input Lock Read Write As #monFichId
    Me.MousePointer = vbDefault
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub BtnZoomIti_Click()
    Dim unT1 As Single, unT2 As Single
    
    'BtnZoomIti.Enabled = False
    'Dessin des axes Distances en Y et temps en X
    DessinerAxes
    'Calcul des temps de passage aux distances correspondant aux rep�res
    'd�but et fin de l'itin�raire charg�
    'Initialisation
    monMaxTReelZIti = -100
    monMinTReelZIti = 10000000
    For i = 1 To maColParcoursMTB.Count
        If maColParcoursMTB(i).monIsUtil Then
            maColParcoursMTB(i).DonnerTemps monMinDReelZIti * 10, monMaxDReelZIti * 10, unT1, unT2
            'conversion des temps des dixi�mes de secondes en minutes
            unT1 = unT1 / 600
            unT2 = unT2 / 600
            If unT1 <= monMinTReelZIti Then monMinTReelZIti = unT1
            If unT2 >= monMaxTReelZIti Then monMaxTReelZIti = unT2
        End If
    Next i
    'Dessin des parcours issus du fichier MTB = campagne de mesures
    DessinerDesParcours ZoneDessin, maColParcoursMTB, maMargeD, monDec, maMargeB, maMargeB, monMinTReelZIti, monMaxTReelZIti, monMinDReelZIti, monMaxDReelZIti, monIndAncienParChoisi
    'Dessin des rep�res itin�raire au zoom tout
    DessinerRepereIti ZoneDessin, maMargeD, monDec, maMargeB, maMargeB, monMinTReelZIti, monMaxTReelZIti, monMinDReelZIti, monMaxDReelZIti
    BtnZoomTout.Enabled = True
    'Stockage des min et max r�els en X et Y
    monMaxYReel = monMaxDReelZIti
    monMinYReel = monMinDReelZIti
    monMaxXReel = monMaxTReelZIti
    monMinXReel = monMinTReelZIti
End Sub

Private Sub BtnZoomTout_Click()
    MousePointer = vbHourglass
    
    'Dessin en zoom tout
    DessinerZoomTout
    'D�sactivation du bouton Zoom tout
    BtnZoomTout.Enabled = False
    'Activation du bouton Zoom iti
    BtnZoomIti.Enabled = True
    'Stockage des min et max r�els en X et Y
    monMaxYReel = monMaxDReelZTout
    monMinYReel = monMinDReelZTout
    monMaxXReel = monMaxTReelZTout
    monMinXReel = monMinTReelZTout
    
    MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    'Pour ne dessiner qu'au premier activate
    If monMaxTReelZTout > 0 And monMaxDReelZTout > 0 Then
        'Apr�s il y a chargement d'un fichier MTB et les deux max sont tjs > 0
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    monIndAncienParChoisi = 0
    Caption = "Import de " + Tag
    
    'Mettre tous les parcours issus du fichier MTB � utiliser pour qu'ils s'affichent
    'For i = 1 To maColParcoursMTB.Count
    '    maColParcoursMTB(i).monIsUtil = True
    'Next i
    
    'Calcul des maxi r�elles en distances et en temps pour le zoom total
    DonnerMaxDistDuree maColParcoursMTB, monMaxDReelZTout, monMaxTReelZTout
    'Calcul des mini r�elles en distances et en temps pour le zoom total
    monMinDReelZTout = 0
    monMinTReelZTout = 0
    
    'Stockage des min et max r�els en X et Y
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
    'Vidage de la collection des indices de parcours import�s
    ViderCollection maColParImport
    'Initialisation du nombre de rep�res cr��s
    'et du nombre de parcours import�s
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
    
    'Placement et taille de la fen�tre
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
    'fichier itin�raire, le nom du parcours � importer
    'et sa date de mesure
    TextFichItiRef.Width = BtnFermer.Left - TextFichItiRef.Left - 75
    
    TextDatePar.Left = BtnFermer.Left - TextDatePar.Width - 75
    LabelDatePar.Left = TextDatePar.Left - LabelDatePar.Width '- 60
    TextNomPar.Width = LabelDatePar.Left - TextNomPar.Left - 60
    TextNomPar.Text = "Aucun"
    
    'Placement de la l�gende en bas � droite
    FrameLegende.Top = ScaleHeight - FrameLegende.Height - 60
    FrameLegende.Left = BtnFermer.Left
    
    'Positionnement de la frame permettant la modif des rep�res d�but
    'et fin d'un itin�raire
    FrameDebFinIti.Left = BtnFermer.Left
    FrameDebFinIti.Top = BtnAide.Top + BtnAide.Height + 120
    FrameDebFinIti.Height = FrameLegende.Top - FrameDebFinIti.Top - 120
    VScrollDebIti.Height = FrameDebFinIti.Height - 600
    VScrollFinIti.Height = FrameDebFinIti.Height - 600
    VScrollDebIti.Enabled = False
    VScrollFinIti.Enabled = False
    LabelCoefEta.Enabled = False
    TextCoefEta.Enabled = False
    'Indication d'une modif des vscrollers rep�res d�but et fin
    'par click souris sur ses fl�ches ==> D�placerTousRepIti fera son travail
    monIsScrollByClick = True
End Sub

Public Sub DessinerZoomTout()
    'Dessiner dans la zone de dessin � un niveau de zoom montrant toutes
    'les entit�s graphiques
    
    'Dessin des axes
    DessinerAxes
    'Dessin des parcours issus du fichier MTB = campagne de mesures
    DessinerDesParcours ZoneDessin, maColParcoursMTB, maMargeD, monDec, maMargeB, maMargeB, monMinTReelZTout, monMaxTReelZTout, monMinDReelZTout, monMaxDReelZTout, monIndAncienParChoisi
    'Dessin des rep�res itin�raire au zoom tout
    DessinerRepereIti ZoneDessin, maMargeD, monDec, maMargeB, maMargeB, monMinTReelZTout, monMaxTReelZTout, monMinDReelZTout, monMaxDReelZTout
End Sub

Public Sub DessinerRepereIti(uneZoneDes As PictureBox, uneMargeD As Single, uneMargeG As Single, uneMargeB As Single, uneMargeH As Single, unMinXreel As Single, unMaxXreel As Single, unMinYreel As Single, unMaxYreel As Single)
    'Dessin des rep�res de l'itin�raire de r�f�rence charg� dans la large droite d'
    'une picture box avec respect des marges entre un min en X et un max en X r�el
    'et entre un min en Y r�el et un max Y r�el
    Dim unX1 As Single, unX2 As Single, unD1 As Single, unD2 As Single
    Dim unXecran As Single, unYecran As Single
    Dim unXecSuiv As Single, unYecSuiv As Single
    Dim unMinXecran As Single, unMaxXecran As Single
    Dim unMinYecran As Single, unMaxYecran As Single
    Dim uneDistMaxReelX As Single, uneDistMaxEcranX As Single
    Dim uneDistMaxReelY As Single, uneDistMaxEcranY As Single
    
    'D�termination des min/max �cran
    'Variables servant pour la conversion coordonn�es r�elles en �cran
    uneDistMaxReelX = unMaxXreel - unMinXreel
    uneDistMaxEcranX = uneZoneDes.Width - uneMargeG - uneMargeD
    uneDistMaxReelY = unMaxYreel - unMinYreel
    uneDistMaxEcranY = uneZoneDes.Height - uneMargeH - uneMargeB

    'Conversion en coordonn�es Y �cran des distances r�elles
    'Les Y sont orient�s vers le bas, donc le max r�el correspondant au max �cran
    'est inf�rieur au min �cran correspondant au min r�el
    unMaxYecran = uneMargeH
    unMinYecran = uneZoneDes.Height - uneMargeB
    
    'Conversion en coordonn�es X �cran des temps ou vitesses r�els
    unMinXecran = uneMargeG
    unMaxXecran = uneZoneDes.Width - uneMargeD
    
    'Calcul de X ecran de positionnement du rond symbolisant le rep�re
    'et la fin de la ligne de rappel
    unXecran = unMaxXecran + monDec
    
    'Dessiner des rep�res et de leur ligne de rappel
    For i = 1 To monSaveNbRepIti
        'R�cup de l'abs curviligne en m�tres
        unD1 = CSng(ShapeRep(i).Tag)
        
        'Placement en Y �cran du rond symbolisant le rep�re et de la ligne de rappel
        'Conversion en coordonn�es �crans des coordonn�es r�elles
        unYecran = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD1, uneDistMaxReelY, uneDistMaxEcranY)
        'Dessin du rep�re et de sa ligne de rappel
        's'il est entre le min et le max y �cran
        'min y �cran  > max y �cran car les y �cran orient�s vers le bas en Y,
        'donc apr�s conversion donn�e r�elle en �cran le max devient < au min
        uneVisu = (unYecran <= unMinYecran + EpsilonEcran And unYecran >= unMaxYecran - EpsilonEcran)
        ShapeRep(i).Visible = uneVisu
        ShapeRep(i).Left = unXecran
        ShapeRep(i).Top = unYecran - TailleRep
        
        LineRep(i).Visible = uneVisu
        LineRep(i).X1 = unMinXecran
        LineRep(i).X2 = unXecran
        LineRep(i).Y1 = unYecran
        LineRep(i).Y2 = unYecran
        
        'Indication du d�but ou fin d'itin�raire
        If i = monIndRepItiDeb Then
            'Cas du rep�re D�but d'itin�raire
            LabelDebIti.Visible = uneVisu
            LabelDebIti.Left = unXecran + ShapeRep(i).Width
            LabelDebIti.Top = unYecran - LabelDebIti.Height / 2
            VScrollDebIti.Tag = 0
            'Indication d'une modif des vscrollers rep�res d�but et fin par programme
            '(.value = qqchose) ==> le change event sur vscroll ne fera rien
            monIsScrollByClick = False
            'Positionnement du vscroll pour le rep�re d�but par rapport
            'au Y �cran du rep�re d�but d'itin�raire
            VScrollDebIti.Value = ConvertirYEcranEnVScrollValue(VScrollDebIti, unYecran)
            'Remise de l'Indication qu'une modif des vscrollers rep�res d�but et fin
            'par click souris sur leurs fl�ches est possible
            '==> le change event sur vscroll fera son travail dans ces cas
            monIsScrollByClick = True
        ElseIf i = monIndRepItiFin Then
            'Cas du rep�re Fin d'itin�raire
            LabelFinIti.Visible = uneVisu
            LabelFinIti.Left = unXecran + ShapeRep(i).Width
            LabelFinIti.Top = unYecran - LabelFinIti.Height / 2
            'Indication d'une modif des vscrollers rep�res d�but et fin par programme
            '(.value = qqchose) ==> le change event sur vscroll ne fera rien
            monIsScrollByClick = False
            'Positionnement du vscroll pour le rep�re fin par rapport
            'au Y �cran du rep�re fin d'itin�raire
            VScrollFinIti.Value = ConvertirYEcranEnVScrollValue(VScrollFinIti, unYecran)
            'Remise de l'Indication qu'une modif des vscrollers rep�res d�but et fin
            'par click souris sur leurs fl�ches est possible
            '==> le change event sur vscroll fera son travail dans ces cas
            monIsScrollByClick = True
            'On grise le vscroll de modif du rep�re fin iti s'il n'est pas visible
            VScrollFinIti.Enabled = uneVisu
            LabelFin.Enabled = uneVisu
        End If
    Next i
End Sub

Private Sub DessinerAxes()
    'Proc�dure dessinant les axes Ox (les temps de parcours)
    'et Oy (les distances parcourues)
    
    monDec = 60
    maMargeB = ZoneDessin.TextHeight("Temps en minutes") + monDec * 2
    maMargeD = 850
    
    'On vide la zone de dessin
    ZoneDessin.Cls
    
    'Axe des temps = Ox
    ZoneDessin.Line (monDec, ZoneDessin.Height - maMargeB)-(ZoneDessin.Width - maMargeD, ZoneDessin.Height - maMargeB), QBColor(0)
    'Dessin de la fl�che
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
    'Dessin de la fl�che
    ZoneDessin.Line (ZoneDessin.Width - maMargeD, monDec)-(ZoneDessin.Width - maMargeD - monDec, monDec * 2), QBColor(0)
    ZoneDessin.Line (ZoneDessin.Width - maMargeD, monDec)-(ZoneDessin.Width - maMargeD + monDec, monDec * 2), QBColor(0)

    'Affichage du texte "Parcours mesur�s" en haut et au milieu de la partie
    'visualisant les parcours issues de la campagne de mesure
    ZoneDessin.CurrentX = (ZoneDessin.Width - maMargeD - monDec * 2 - ZoneDessin.TextWidth("Parcours mesur�s")) / 2
    ZoneDessin.CurrentY = 0
    ZoneDessin.Print "Parcours mesur�s"
    
    'Affichage du texte "Itin�raire" en haut et au milieu de la partie
    'visualisant les rep�res de l'itin�raire de r�f�rence charg�
    ZoneDessin.CurrentX = ZoneDessin.Width - maMargeD + monDec * 2
    ZoneDessin.CurrentY = 0
    ZoneDessin.Print "Itin�raire"
    
    'Dessin du cadre de dessin possible pour tester
    'ZoneDessin.Line (monDec, maMargeB)-(ZoneDessin.Width - maMargeD, ZoneDessin.Height - maMargeB), QBColor(0), B
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Fermeture du fichier itin�raire charg� �ventuel
    If monFichId > 0 Then Close #monFichId
    'Fermeture de la fen�tre de choix lors d'une s�lection de plusieurs parcours
    'possibles
    Unload frmChoixPar
End Sub


Private Sub LabelDebIti_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'S�lection du rep�re d�but d'iti impossible
    MsgBox "La s�lection de rep�re d�but d'itin�raire est impossible.", vbExclamation
End Sub

Private Sub LabelFinIti_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'S�lection du rep�re fin d'iti impossible
    MsgBox "La s�lection de rep�re fin d'itin�raire est impossible.", vbExclamation
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
    'Fonction retournant la coordonn�es en Y �cran donn�e par le champ
    'value des scroller vertical li�s au d�but et fin d'itin�raire charg�
    Dim unMaxYecran As Double, unMinYecran As Double
    Dim unA As Double, unB As Double
    
    'Le max y �cran correpondant au max y r�el est en haut de fen�tre
    'c'est l'inverse pour le min y �cran car Oy r�el est orient� de bas en haut
    'alors que Oy �cran est orient� de haut en bas
    unMaxYecran = maMargeB
    unMinYecran = ZoneDessin.Height - maMargeB
    
    'On veut que unMaxYecran = unA * unVScroll.Min + unB
    'et unMin = unA * unVScroll.Max + unB d'o� le calcul de unA et unB
    unA = (unMaxYecran - unMinYecran) / (unVScroll.Min - unVScroll.Max)
    unB = unMaxYecran - unA * unVScroll.Min
    
    'Calcul de la position en Y �cran
    ConvertirVScrollValueEnYEcran = unA * uneVal + unB
End Function

Private Function ConvertirYEcranEnVScrollValue(unVScroll As VScrollBar, unY As Single) As Single
    'Fonction retournant la value des scroller vertical li�s au d�but
    'et fin d'itin�raire charg�, en fonction de la coordonn�e Y �cran
    'du rep�re d�but ou fin
    Dim unMaxYecran As Single, unMinYecran As Single
    Dim unA As Single, unB As Single, uneVal As Single
    
    'Le max y �cran correpondant au max y r�el est en haut de fen�tre
    'c'est l'inverse pour le min y �cran car Oy r�el est orient� de bas en haut
    'alors que Oy �cran est orient� de haut en bas
    unMaxYecran = maMargeB
    unMinYecran = ZoneDessin.Height - maMargeB
    
    'On veut que unMaxYecran = unA * unVScroll.Min + unB
    'et unMin = unA * unVScroll.Max + unB d'o� le calcul de unA et unB
    unA = (unMaxYecran - unMinYecran) / (unVScroll.Min - unVScroll.Max)
    unB = unMaxYecran - unA * unVScroll.Min
    
    'Calcul de la position en Y �cran
    uneVal = (unY - unB) / unA
    If uneVal < unVScroll.Min Then
        'Cas o� on est en dessous du min du vscroll ===> on s'y met
        ConvertirYEcranEnVScrollValue = unVScroll.Min
    ElseIf uneVal > unVScroll.Max Then
        'Cas o� on d�passe le max du vscroll ===> on s'y met
        ConvertirYEcranEnVScrollValue = unVScroll.Max
    Else
        'Autres cas uneVal entre le min et le max
        ConvertirYEcranEnVScrollValue = uneVal
    End If
End Function


Public Sub DeplacerTousRepIti()
    'Fonction d�placement tous les rep�res itin�raires, donc les ronds les symbolisant
    'et leur ligne de rappel respective et les labels d�but et fin
    'lorsque l'on clique sur l'ascenseur d�pla�ant le rep�re d�but
    Dim uneTransEcran As Single, uneTransReel As Single, unYreel As Single
    Dim uneVisu As Boolean
    Dim unYEcranRepDebAvant As Double, unNewYEcranRepDeb As Double
    Dim uneDistMaxReel As Single, uneDistMaxEcran As Single
    Dim unMinYecran As Single, unMaxYecran As Single
    
    'Indication d'une modif du vscroll rep�re d�but par programme (.value = qqchose)
    '==> D�placerTousRepIti ne fera rien
    If monIsScrollByClick = False Then Exit Sub
    
    'D�termination des min/max �cran
    'Variables servant pour la conversion abscisses curvilignes r�elles en �cran
    unMaxYecran = maMargeB
    unMinYecran = ZoneDessin.Height - maMargeB
    uneDistMaxReel = monMaxYReel - monMinYReel
    uneDistMaxEcran = ZoneDessin.Height - maMargeB * 2
    'Stockage du Y �cran du rep�re d�but avant d�placement
    unYEcranRepDebAvant = LineRep(monIndRepItiDeb).Y1
    'Calcul de la nouvelle position du rep�re d�but
    unNewYEcranRepDeb = ConvertirVScrollValueEnYEcran(VScrollDebIti, VScrollDebIti.Value)
    LineRep(monIndRepItiDeb).Y1 = unNewYEcranRepDeb
    LabelDebIti.Top = unNewYEcranRepDeb - LabelDebIti.Height / 2
    'Calcul de la translation effectu�
    uneTransEcran = unNewYEcranRepDeb - unYEcranRepDebAvant
    'Propagation � tous les rep�res itin�raires et aux controls les symbolisant
    LabelDebIti.Top = LabelDebIti.Top + uneTransEcran
    For i = 1 To monNbRepIti
        ShapeRep(i).Top = ShapeRep(i).Top + uneTransEcran
        If i = monIndRepItiDeb Then
            'On voit toujours le rep�re d�but
            'et son LineRep.Y1 a �t� modifi� ci-dessus
            uneVisu = True
        Else
            LineRep(i).Y1 = LineRep(i).Y1 + uneTransEcran
            uneVisu = (LineRep(i).Y1 <= unMinYecran + EpsilonEcran And LineRep(i).Y1 >= unMaxYecran - EpsilonEcran)
        End If
        LineRep(i).Y2 = LineRep(i).Y1
        ShapeRep(i).Visible = uneVisu
        LineRep(i).Visible = uneVisu
        'Modification de l'abscisse curviligne r�el stock� dans le tag
        'en convertissant l'abs curv �cran en abs curv r�elle
        unYreel = ConvertirEnReel(monMaxYReel, maMargeB - LineRep(i).Y1, uneDistMaxReel, uneDistMaxEcran)
        ShapeRep(i).Tag = Format(unYreel)
    Next i
    'D�placement du label Fin du rep�re fin
    LabelFinIti.Top = LineRep(monIndRepItiFin).Y1 - LabelFinIti.Height / 2
    'Affichage des labels d�but/fin si la ligne de rappel du rep�re d�but/fin
    'est visible
    LabelDebIti.Visible = LineRep(monIndRepItiDeb).Visible
    LabelFinIti.Visible = LineRep(monIndRepItiFin).Visible
    'On grise le vscroll de modif du rep�re fin iti s'il n'est pas visible
    VScrollFinIti.Enabled = LineRep(monIndRepItiFin).Visible
    LabelFin.Enabled = LineRep(monIndRepItiFin).Visible
    'D�placement des min et max r�els de l'itin�raire d'une valeur de
    'translation r�el en convertissant la translation �cran
    uneTransReel = DonnerDistReel(uneTransEcran, uneDistMaxReel, uneDistMaxEcran)
    monMaxDReelZIti = monMaxDReelZIti - uneTransReel
    monMinDReelZIti = monMinDReelZIti - uneTransReel

    'Mise � jour de la position du scroller modifiant le rep�re iti r�f de fin
    'Synchronisation de la nouvelle position dans le vscroll fin
    'Indication d'une modif des vscrollers rep�res d�but et fin par programme
    '(.value = qqchose)==> le change event sur vscroll ne fera rien
    monIsScrollByClick = False
    'Positionnement du vscroll pour le rep�re fin par rapport
    'au Y �cran du rep�re fin d'itin�raire
    VScrollFinIti.Value = ConvertirYEcranEnVScrollValue(VScrollDebIti, LineRep(monIndRepItiFin).Y1)
    'Remise de l'Indication qu'une modif des vscrollers rep�res d�but et fin
    'par click souris sur leurs fl�ches est possible
    '==> le change event sur vscroll fera son travail dans ces cas
    monIsScrollByClick = True
End Sub

Public Sub EtirerTousRepIti()
    'Fonction �tirant tous les rep�res itin�raires, donc les ronds les symbolisant
    'et leur ligne de rappel respective et les labels d�but et fin
    'lorsque l'on clique sur l'ascenseur d�pla�ant le rep�re fin
    'donc il peut y avoir contraction ou dilatation des abs curv ou y des rep�res
    Dim unCoefEtirageEcran As Double, unCoefEtirageReel As Single
    Dim unYreel As Single, unYEcranRepDeb As Double
    Dim unYEcranRepFinAvant As Double, unNewYEcranRepFin As Double
    Dim uneDistMaxReel As Single, uneDistMaxEcran As Single
    
    'Indication d'une modif du vscroll rep�re d�but par programme (.value = qqchose)
    '==> D�placerTousRepIti ne fera rien
    If monIsScrollByClick = False Then Exit Sub
    
    'Stockage du Y �cran du rep�re d�but
    unYEcranRepDeb = LineRep(monIndRepItiDeb).Y1
    'D�termination des min/max �cran
    'Variables servant pour la conversion abscisses curvilignes r�elles en �cran
    uneDistMaxReel = monMaxYReel - monMinYReel
    uneDistMaxEcran = ZoneDessin.Height - maMargeB * 2
    'Stockage du Y �cran du rep�re fin avant d�placement
    unYEcranRepFinAvant = LineRep(monIndRepItiFin).Y1
    'Calcul de la nouvelle position du rep�re fin
    unNewYEcranRepFin = ConvertirVScrollValueEnYEcran(VScrollFinIti, VScrollFinIti.Value)
    'LineRep(monIndRepItiFin).Y1 = unNewYEcranRepFin
    'Calcul de l'�tirage effectu�
    unCoefEtirageEcran = (unNewYEcranRepFin - unYEcranRepDeb) / (unYEcranRepFinAvant - unYEcranRepDeb)
    'Affichage du coefficient d'�tirage = coef d'�talonnage
    TextCoefEta.Text = Format(CSng(TextCoefEta.Text) * unCoefEtirageEcran, "##0.0000")
    'Propagation � tous les rep�res itin�raires et aux controls les symbolisant
    'sauf le rep�re d�but d'itin�raire
    For i = 1 To monNbRepIti
        If i <> monIndRepItiDeb Then
            LineRep(i).Y1 = unYEcranRepDeb + (LineRep(i).Y1 - unYEcranRepDeb) * unCoefEtirageEcran
            LineRep(i).Y2 = LineRep(i).Y1
            ShapeRep(i).Top = LineRep(i).Y1 - ShapeRep(i).Height / 2
            'Modification de l'abscisse curviligne r�el stock� dans le tag
            'en convertissant l'abs curv �cran en abs curv r�elle
            unYreel = ConvertirEnReel(monMaxYReel, maMargeB - LineRep(i).Y1, uneDistMaxReel, uneDistMaxEcran)
            ShapeRep(i).Tag = Format(unYreel)
        End If
    Next i
    'Mise au premier plan du rep�re fin
    ShapeRep(monIndRepItiFin).ZOrder 0
    'D�placement du label fin du rep�re fin
    LabelFinIti.Top = unNewYEcranRepFin - LabelFinIti.Height / 2
    'D�placement du max r�el de l'itin�raire en convertissant la coordonn�e �cran
    'du rep�re fin en coordonn�e relle
    monMaxDReelZIti = ConvertirEnReel(monMaxYReel, maMargeB - LineRep(monIndRepItiFin).Y1, uneDistMaxReel, uneDistMaxEcran)
End Sub

Private Sub ZoneDessin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'S�lection des rep�res itin�raires ou des rep�res de parcours
    'Pour la s�lection des rep�res itin�raires, on alimente la variable globale
    'priv�e monIndRepItiSel donnant l'indice du control ShapeRep cliqu�,
    '0 si rien de s�lectionner
    Dim unPar As Parcours, unD1 As Single, unX1 As Single, unY1 As Single
    Dim unX2 As Single, unY2 As Single, uneDist As Single, j As Long
    Dim unX As Single, unY As Single
    Dim uneDistMaxReelX As Single, uneDistMaxEcranX As Single
    Dim uneDistMaxReelY As Single, uneDistMaxEcranY As Single
    Dim unIndParChoisi As Integer
    
    'Si aucun itin�raire charg�, aucun s�lection n'est possible
    'Le bouton changer d�but itin�raire apparait d�s le premier chargement d'itin�raire
    If BtnChgDebIti.Enabled = False Then
        MsgBox "Les s�lections graphiques des parcours sont possibles uniquement si un itin�raire est charg�.", vbInformation
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    
    'D�termination des min/max �cran
    'Variables servant pour la conversion abscisses curvilignes r�elles en �cran
    uneDistMaxReelY = monMaxYReel - monMinYReel
    uneDistMaxEcranY = ZoneDessin.Height - maMargeB * 2
    uneDistMaxReelX = monMaxXReel - monMinXReel
    uneDistMaxEcranX = ZoneDessin.Width - monDec - maMargeD
    
    'Initialisation
    i = 0
    unIndParChoisi = 0
    'Vidage de la listbox des parcours trouv�s
    frmChoixPar.Visible = False
    frmChoixPar.ListParTrouv.Clear
    TextNomPar.Text = "Aucun"
    TextDatePar.Text = ""
    
    If X >= ZoneDessin.Width - maMargeD + monDec And X <= ZoneDessin.Width Then
        'Cas o� l'on clique � droite de l'axe Oy, il n'y a que des rep�res itin�raires
        'Recherche du rond , donc du rep�re itin�raire cliqu�
        For i = 1 To monNbRepIti
            If X - ShapeRep(i).Left >= 0 And X - ShapeRep(i).Left <= ShapeRep(i).Width Then
                If Y - ShapeRep(i).Top >= 0 And Y - ShapeRep(i).Top <= ShapeRep(i).Height Then
                    'Cas o� un rep iti est trouv�
                    If i = monIndRepItiDeb Or i = monIndRepItiFin Then
                        'S�lection du rep�re d�but ou fin d'iti impossible
                        MsgBox "La s�lection de rep�re d�but ou fin d'itin�raire est impossible.", vbExclamation
                    Else
                        'D�selection de l'ancienne s�lection, bordure n'est plus noire
                        ShapeRep(monIndRepItiSel).BorderColor = ShapeRep(monIndRepItiSel).BackColor
                        'Stockage de l'index de s�lection et s�lection,
                        'la bordure est mise en noir autour du rep iti trouv�
                        ShapeRep(i).BorderColor = QBColor(0)
                        monIndRepItiSel = i
                    End If
                    Exit For
                End If
            End If
        Next i
    Else
        'Autres cas, on est dans une partie o� seul un parcours
        'issu du fichier MTB peut �tre cliqu�
        i = 0 'pour indiquer de faire la d�selection des rep iti
        
        'Conversion d'une distance de EpsilonEcran/2 twips en distance r�elle sur OY et du Y �cran
        unEpsilonY = DonnerDistReel(EpsilonEcran / 2, monMaxYReel - monMinYReel, uneDistMaxEcranY)
        unY = ConvertirEnReel(monMinYReel, ZoneDessin.Height - maMargeB - Y, uneDistMaxReelY, uneDistMaxEcranY)
        'Conversion d'une distance de EpsilonEcran/2 twips en distance r�elle sur OY
        unEpsilonX = DonnerDistReel(EpsilonEcran / 2, monMaxXReel - monMinXReel, uneDistMaxEcranX)
        unX = ConvertirEnReel(monMinXReel, X - monDec, uneDistMaxReelX, uneDistMaxEcranX)
        'Epsilon pour la proximit� en diagonale ou en projection
        unEpsilonXY = Sqr(unEpsilonX * unEpsilonX + unEpsilonY * unEpsilonY)
        
        'Recherche sur tous les parcours affich�s du parcours cliqu�
        For k = 1 To maColParcoursMTB.Count
            Set unPar = maColParcoursMTB(k)
            If unPar.monIsUtil Then
                unNbPoints = unPar.monNbPas
                'Conversion dixi�me de secondes en minutes
                unX1 = unPar.monFirstPas / 600
                uneDist = 0
                    
                For j = 2 To unNbPoints
                    'Recup des coordonn�es x des points
                    'Conversion des dixi�mes de seconde et des secondes en minutes
                    If j = unNbPoints Then
                        unX2 = unPar.monFirstPas / 600 + (unNbPoints - 2) * unPar.monPasMesure / 60 + unPar.monLastPas / 600
                    Else
                        unX2 = unPar.monFirstPas / 600 + (j - 1) * unPar.monPasMesure / 60
                    End If
                                    
                    'Calcul du Y
                    uneDist = uneDist + unPar.monTabDist(j - 1) * unPar.monCoefEta / 10
                    unY1 = uneDist 'Stockage pour l'incr�mentation suivante
                    unY2 = uneDist + unPar.monTabDist(j) * unPar.monCoefEta / 10
                    
                    'Si point confondu on ne fait rien, on passe au suivant
                    unPtConfondu = (unX1 = unX2 And unY1 = unY2)
                    If unPtConfondu = False Then
                        'Recherche si on a cliqu� pr�s du segment M1(x1,y1)-M2(x2,y2)
                        'D'abord on regarde si X1 < X < X2, puis Y1 < Y < Y2 et enfin
                        'si la distance � la droite passant par M1 et M2 est < espilon
                        If (unX1 - unEpsilonX < unX) And (unX < unX2 + unEpsilonX) Then
                            If (unY1 - unEpsilonY < unY) And (unY < unY2 + unEpsilonY) Then
                                'Calcul de la distance � la droite pasant par M1M2
                                'ax+by+c = 0, a = y2-y1, b=x1-x2, c = x2y1-x1y2
                                uneDistM1M2 = (unY2 - unY1) * unX + (unX1 - unX2) * unY + (unX2 * unY1 - unX1 * unY2)
                                If Abs(uneDistM1M2) / Sqr((unX2 - unX1) * (unX2 - unX1) + (unY2 - unY1) * (unY2 - unY1)) < unEpsilonXY Then
                                    'Cas o� le parcours est clqu�
                                    'Stockage du parcours et des x et y �cran du rep�re s�lectionn�
                                    unIndParChoisi = k
                                    'Ajout dans la listbox des parcours trouv�s
                                    frmChoixPar.ListParTrouv.AddItem unPar.monNom + " (" + Mid(unPar.monJourSemaine, 1, 2) + " " + Format(unPar.maDate) + " " + Mid(Format(unPar.monHeureDebut), 1, 5) + ")"
                                    frmChoixPar.ListParTrouv.ItemData(frmChoixPar.ListParTrouv.NewIndex) = k
                                    'Sortie du for bouclant sur les points du parcours
                                    'on passe au parcours suivant
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                    'Pour l'incr�mentation suivante
                    unX1 = unX2
                Next j
            End If
        Next k
        If unIndParChoisi = 0 Then
            MsgBox "Aucun parcours n'a �t� s�lectionn�.", vbInformation
        End If
    End If
    
    'Si aucun s�lection de rep�re, on d�selectionne le dernier rep iti s�lectionn�
    If i = 0 Or i = monNbRepIti + 1 Then
        ShapeRep(monIndRepItiSel).BorderColor = ShapeRep(monIndRepItiSel).BackColor
        monIndRepItiSel = 0
    End If
    
    'Si plus d'un parcours s�lectionn�, on affichage leur nom dans une fen�tre
    'de choix qui apparait au centre de l'�cran
    If frmChoixPar.ListParTrouv.ListCount > 1 Then
        CentrerFenetreEcran frmChoixPar
        frmChoixPar.Show vbModal
        If Tag = "" Then
            'Cas o� aucun parcours choisi
            '===> click sur bouton Annuler de la fen�tre choix parcours
            unIndParChoisi = 0
        Else
            'Cas o� un parcours  a �t� choisi
            unIndParChoisi = CInt(Tag)
        End If
    End If
            
    If unIndParChoisi > 0 Then
        'Cas o� un parcours a �t� choisi par click sur un rep�re
        'ou par choix parmi plusieurs et on l'affiche en trait �pais
        TextNomPar.Text = maColParcoursMTB(unIndParChoisi).monNom
        TextDatePar.Text = maColParcoursMTB(unIndParChoisi).monJourSemaine + " " + Format(maColParcoursMTB(unIndParChoisi).maDate) + " " + Format(maColParcoursMTB(unIndParChoisi).monHeureDebut)
        If frmChoixPar.ListParTrouv.ListCount = 1 Then
            'On redessine le parcours choisi en trait gros uniquement
            'si un seul parcours trouv�, sinon on redessine plusieurs fois
            MontrerParcoursChoisi unIndParChoisi
        End If
        'Stockage de l'indice du parcours s�lection
        monIndAncienParChoisi = unIndParChoisi
    ElseIf monIndAncienParChoisi > 0 Then
        'Cas o� rien de s�lectionner par click, mais o� un parcours
        'avait �t� cliqu� avant, on met l'ancienne s�lection � 0 pour
        'tout redessiner en trait fin
        monIndAncienParChoisi = 0
        MontrerParcoursChoisi monIndAncienParChoisi
    ElseIf Mid(monBtnClick, 1, 7) = "Annuler" And Len(monBtnClick) > 7 Then
        'Longueur > 7 car monBtnclcik peut ne valoir qu'annuler si on a utilis�
        'd'autres fen�tres avant qui avait des boutons ok et annuler remplissant
        'la variable globale monBtnClick, frmChoixPar donne Annuler plus l'index du
        'parcours s�lectionn� ou -1 si aucun parcours n'est s�lectionn�
        If CInt(Mid(monBtnClick, 8)) > -1 Then
            'On redessine tous les parcours choisi en trait fin en disant qu'aucun n'a
            '�t� trouv�
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
    'Dessin des parcours en mettant en trait �pais le parcours choisi
    DessinerDesParcours ZoneDessin, maColParcoursMTB, maMargeD, monDec, maMargeB, maMargeB, monMinXReel, monMaxXReel, monMinYReel, monMaxYReel, unIndParChoisi
End Sub
