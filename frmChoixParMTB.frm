VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmChoixParMTB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choix des parcours issus du fichier MTB à importer"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   Icon            =   "frmChoixParMTB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread SpreadParcoursMTB 
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "frmChoixParMTB.frx":0442
      TabIndex        =   0
      Top             =   120
      Width           =   10935
   End
   Begin VB.CommandButton btnDeselTout 
      Caption         =   "Désélectionner tout"
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
      Left            =   5640
      TabIndex        =   4
      Top             =   5040
      Width           =   2647
   End
   Begin VB.CommandButton btnSelTout 
      Caption         =   "Sélectionner tout"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   5040
      Width           =   2647
   End
   Begin VB.CommandButton btnVoirParSelect 
      Caption         =   "Visualiser les parcours sélectionnés"
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
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   2647
   End
   Begin VB.CommandButton btnAnnuler 
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
      Left            =   8400
      TabIndex        =   1
      Top             =   5040
      Width           =   2647
   End
End
Attribute VB_Name = "frmChoixParMTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAnnuler_Click()
    'On vide les parcours issus du mtb
    ViderColParcours maColParcoursMTB
    
    'Fermeture de la fenêtre et stockage du fait que
    'l'on a cliqué sur le bouton Annuler
    Tag = ""
    monBtnClick = "Annuler"
    Unload Me
End Sub

Private Sub btnDeselTout_Click()
    DecocherToutColSelection SpreadParcoursMTB
End Sub

Private Sub btnSelTout_Click()
    CocherToutColSelection SpreadParcoursMTB
End Sub

Private Sub btnVoirParSelect_Click()
    Dim unNbParSel As Integer, i As Integer
    
    'Initialisation
    unNbParSel = 0
    
    'Boucle sur les  parcours trouvés dans le fichier MTB pour voir
    'si au moins un parcours est sélectionné
    For i = 1 To maColParcoursMTB.Count
        Set unParcours = maColParcoursMTB(i)
        SpreadParcoursMTB.Row = i
        SpreadParcoursMTB.Col = 2
        'Test si le parcours est sélectionné
        ' value = 1 ==> case cochée, 0 sinon
        If SpreadParcoursMTB.Value = 1 Then
            unNbParSel = unNbParSel + 1
        End If
    Next i
    
    If unNbParSel = 0 Then
        'Cas où aucun parcours sélectionné
        MsgBox "Vous devez sélectionner au moins un parcours, en cochant l'une des cases de la colonne sélection.", vbExclamation
    Else
        'Boucle sur les parcours en ne s'occupant que des parcours
        'sélectionnés et on met le champ util à vrai,
        'Les autres non sélectionnés seront à faux et donc
        'on ne le verra pas dans la fenêtre graphique d'import
        'frmImportMTB
        For i = 1 To maColParcoursMTB.Count
            Set unParcours = maColParcoursMTB(i)
            SpreadParcoursMTB.Row = i
            SpreadParcoursMTB.Col = 2
            'Test si le parcours est sélectionné
            ' value = 1 ==> case cochée, 0 sinon
            If SpreadParcoursMTB.Value = 1 Then
                'Indication que le parcours i va être utilisé et
                'affiché dans la fenêtre graphique d'import
                unParcours.monIsUtil = True
            Else
                'Indication que le parcours i ne va pas être utilisé et
                'affiché dans la fenêtre graphique d'import
                unParcours.monIsUtil = False
            End If
        Next i
        'Fermeture de la fenêtre et stockage du fait que
        'l'on a cliqué sur le bouton Visualisation = OK
        monBtnClick = "OK"
        Unload Me
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        'Cas où l'on tape la touche Echap
        If SpreadParcoursMTB.ActiveCol <> 6 Or SpreadParcoursMTB.EditMode = False Then
            'On fait le traitement du bouton annuler
            'sinon on laisse l'annulation de la frappe en
            'cours dans la saisie du coef d'étalonnage
            btnAnnuler_Click
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim unParcours As Parcours
    Dim unNbRepTop As Integer
    
    SpreadParcoursMTB.LockBackColor = vbInfoBackground
    CentrerFenetreEcran Me
    MousePointer = vbHourglass
    
    'Contexte d'aide
    HelpContextID = HelpID_WinImportMesure
    
    'Affichage d'un libellé au coin haut gauche
    SpreadParcoursMTB.Row = 0
    SpreadParcoursMTB.Col = 0
    SpreadParcoursMTB.Text = "N°"
    
    'Initialisation du nombre de lignes du spread
    SpreadParcoursMTB.MaxRows = maColParcoursMTB.Count
    'Initialisation de la colonne du coef d'étalonnage du spread
    'avec le séparateur décimal en cours
    TrouverCaractèreDécimalUtilisé
    InitColSpreadCaractèreDécimal SpreadParcoursMTB, 6, monCarDeci
    'Affichage du contenu de la collection des parcours du MTB
    'dans le spread de cette fenêtre
    For i = 1 To maColParcoursMTB.Count
        Set unParcours = maColParcoursMTB(i)
        SpreadParcoursMTB.Row = i
        SpreadParcoursMTB.Col = 1
        SpreadParcoursMTB.Text = unParcours.monNom
        SpreadParcoursMTB.Col = 2
        SpreadParcoursMTB.Value = 1 'pour être sélectionné Abs(unParcours.monIsUtil)
        'En effet value = 1 si case cochée, 0 sinon
        'true = -1 et false = 0 et ICI sélection = utilsé
        SpreadParcoursMTB.Col = 3
        SpreadParcoursMTB.Text = unParcours.monJourSemaine
        SpreadParcoursMTB.Col = 4
        SpreadParcoursMTB.Text = unParcours.maDate
        SpreadParcoursMTB.Col = 5
        SpreadParcoursMTB.Text = unParcours.monHeureDebut
        SpreadParcoursMTB.Col = 6
        SpreadParcoursMTB.Text = unParcours.monCoefEta
        SpreadParcoursMTB.Col = 7
        'Récup de la distance au dernier top = abs curviligne du dernier top
        unNbRepTop = UBound(unParcours.monTabAbsRep)
        SpreadParcoursMTB.Text = CLng(unParcours.monTabAbsRep(unNbRepTop) * unParcours.monCoefEta / 10)
        SpreadParcoursMTB.Col = 8
        SpreadParcoursMTB.Text = unNbRepTop
    Next i
    MousePointer = vbDefault
End Sub

Private Sub SpreadParcoursMTB_KeyUp(KeyCode As Integer, Shift As Integer)
'Private Sub SpreadParcoursMTB_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    Dim unPar As Parcours
    
    If SpreadParcoursMTB.ActiveCol = 6 Then
        'Cas de la saisie du coefficient d'étalonnage
        SpreadParcoursMTB.Row = SpreadParcoursMTB.ActiveRow
        SpreadParcoursMTB.Col = SpreadParcoursMTB.ActiveCol
        Set unPar = maColParcoursMTB(SpreadParcoursMTB.Row)
        unPar.monCoefEta = Format(SpreadParcoursMTB.Text)
        'Modif de la distance du dernier top
        SpreadParcoursMTB.Col = 7
        'Distance au dernier top = abs curviligne du dernier top
        unNbRepTop = UBound(unPar.monTabAbsRep)
        SpreadParcoursMTB.Text = CLng(unPar.monTabAbsRep(unNbRepTop) * unPar.monCoefEta / 10)
    End If
End Sub

Private Sub SpreadParcoursMTB_KeyPress(KeyAscii As Integer)
    If SpreadParcoursMTB.ActiveCol = 6 Then
        'Cas de la saisie du coefficient d'étalonnage
        'Récup du séparateur décimale en cours
        TrouverCaractèreDécimalUtilisé
        SpreadParcoursMTB.Row = SpreadParcoursMTB.ActiveRow
        SpreadParcoursMTB.Col = SpreadParcoursMTB.ActiveCol
        If monCarDeci = "." And KeyAscii = 44 Then
            'on remplace la virgule (ascii = 44) par le séparateur
            'décimale en cours pour utiliser le clavier numérique
            KeyAscii = Asc(monCarDeci)
        ElseIf monCarDeci = "," And KeyAscii = 46 Then
            'on remplace le point (ascii = 46) par le séparateur
            'décimale en cours pour utiliser le clavier numérique
            KeyAscii = Asc(monCarDeci)
        End If
    End If
End Sub
