VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmNewParMTB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cr�ation d'un itin�raire de r�f�rence � partir des parcours du fichier MTB"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   Icon            =   "frmNewParMTB.frx":0000
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
      OleObjectBlob   =   "frmNewParMTB.frx":0442
      TabIndex        =   0
      Top             =   120
      Width           =   10935
   End
   Begin VB.CommandButton btnDeselTout 
      Caption         =   "D�s�lectionner tout"
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
      Caption         =   "S�lectionner tout"
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
   Begin VB.CommandButton btnCreerItiRef 
      Caption         =   "Cr�er un itin�raire � partir des parcours s�lectionn�s"
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
Attribute VB_Name = "frmNewParMTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable stockant pour chaque parcours la liste des �carts
'entre les abs curv des rep�res et la valeur moyenne
Public maColEcart As New Collection

'Variable stockant le nombre de parcours s�lectionn�s
Public monNbParSel As Integer

Private Sub btnAnnuler_Click()
    'On vide les parcours issus du mtb
    ViderColParcours maColParcoursMTB
    
    'Fermeture de la fen�tre
    Unload Me
End Sub

Private Sub btnCreerItiRef_Click()
    Dim uneSortie As Boolean, i As Integer
    Dim unNbRep0 As Integer, unNbRep As Integer
    Dim unCoefEta As Single, uneDistLastTop As Single
    Dim unMsg As String, uneDistMoyFinal As Single
    Dim uneDist As Single, unParcours As Parcours
    Dim unEcart  As Single
    
    'Initialisation
    ViderCollection maColValMoy
    ViderCollection maColEcart
    monNbParSel = 0
    unNbRep0 = -1
    uneSortie = False
    i = 1
    
    'Boucle sur les  parcours trouv�s pour voir si au moins est s�lectionn�
    'et v�rifier si tous les parcours s�lectionn�s ont le m�me nombre de
    'rep�res et calcul de la distance moyenne au dernier top des parcours
    's�lectionn�s
    While i > 0 And i < maColParcoursMTB.Count + 1 And uneSortie = False
        Set unParcours = maColParcoursMTB(i)
        SpreadParcoursMTB.Row = i
        SpreadParcoursMTB.Col = 2
        'Test si le parcours est s�lectionn�
        ' value = 1 ==> case coch�e, 0 sinon
        If SpreadParcoursMTB.Value = 1 Then
            'R�cup du nombre de rep�res
            unNbRep = UBound(unParcours.monTabAbsRep)
            If unNbRep0 = -1 Then
                'Initialisation de unNbRep avec le nb de rep�res du premier
                'parcours s�lectionn�
                unNbRep0 = unNbRep
                'Stockage pour calcul de la distance moyenne au dernier top
                monNbParSel = monNbParSel + 1
                SpreadParcoursMTB.Col = 6 'R�cup du coef d'�talonnage
                unCoefEta = CSng(SpreadParcoursMTB.Text)
                'Ajout dans la collection des valeurs moyennes
                'd'abs curvilignes de rep�res et de distance parcourue en fin
                'et de la dur�e de mesure
                For j = 1 To unNbRep0
                    maColValMoy.Add unParcours.monTabAbsRep(j) / 10 * unCoefEta
                Next j
            ElseIf unNbRep0 <> unNbRep Then
                'Cas o� les nb de rep�res sont diff�rents ==> on sort
                uneSortie = True
            Else
                'Stockage pour calcul de la distance moyenne au dernier top
                monNbParSel = monNbParSel + 1
                SpreadParcoursMTB.Col = 6 'R�cup du coef d'�talonnage
                unCoefEta = CSng(SpreadParcoursMTB.Text)
                'Cumul dans la collection des valeurs moyennes
                'd'abs curvilignes de rep�res et de distance parcourue en fin
                For j = 1 To unNbRep0
                    'Insertion avant du nouveau cumul
                    maColValMoy.Add maColValMoy(j) + unParcours.monTabAbsRep(j) / 10 * unCoefEta, , j
                    'Suppression ancien cumul qui s'est d�cal� en j+1
                    maColValMoy.Remove j + 1
                Next j
            End If
        End If
        i = i + 1
    Wend
    
    If unNbRep0 = -1 Then
        'Cas o� aucun parcours s�lectionn� car unNbRep non modifi�, tjs = -1
        ViderCollection maColValMoy
        MsgBox "Vous devez s�lectionner au moins un parcours, en cochant l'une des cases de la colonne s�lection.", vbExclamation
    ElseIf uneSortie = True Then
        'Cas o� les nb de rep�res sont diff�rents
        ViderCollection maColValMoy
        MsgBox "Vous devez s�lectionner des parcours ayant le m�me nombre de rep�res.", vbExclamation
    Else
        If monNbParSel > 1 Then
            'Calcul des abs curvilignes moyen des rep�res et de
            'la distance parcourue moyenne et de la dur�e moyenne
            'des parcours s�lectionn�s
            For i = 1 To unNbRep0
                'Insertion avant de la nouvelle valeur
                maColValMoy.Add maColValMoy(i) / monNbParSel, , i
                'Suppression ancienne valeur d�cal� en i+1
                maColValMoy.Remove i + 1
            Next i
        End If
        
        'Boucle sur les parcours en ne s'occupant que des parcours
        's�lectionn�s et on affichera une fen�tre d'info d�s qu'un �cart
        'entre une abs curv de rep�re ou la distance parcourue et sa valeur
        'moyenne des parcours s�lectionn�s est inf�rieure � la tol�rance
        'fix�e dans les options de MiTemps
        For i = 1 To maColParcoursMTB.Count
            Set unParcours = maColParcoursMTB(i)
            SpreadParcoursMTB.Row = i
            SpreadParcoursMTB.Col = 2
            'Test si le parcours est s�lectionn�
            ' value = 1 ==> case coch�e, 0 sinon
            If SpreadParcoursMTB.Value = 1 Then
                'Indication que le parcours i va �tre utilis� et
                'affect� dans le nouvel itin�raire
                unParcours.monIsUtil = True
                'R�cup du coef d'�talonnage
                SpreadParcoursMTB.Col = 6
                unCoefEta = CSng(SpreadParcoursMTB.Text)
                'Test des �carts des abs curv rep�res avec les valeurs
                'moyennes par rapport � la tol�rance des options logiciels
                For j = 1 To unNbRep0
                    uneDist = unParcours.monTabAbsRep(j) / 10 * unCoefEta
                    'Calcul de l'�cart � la moyenne en %
                    If maColValMoy(j) = 0 Then
                        'Cas o� la moyenne de nombre > 0 vaut 0
                        '==> tous les nombres = 0
                        '==> Ecart = 0
                        unEcart = 0
                    Else
                        unEcart = Abs(maColValMoy(j) - uneDist) / maColValMoy(j) * 100
                    End If
                    'On ajoute � la liste des �carts
                    maColEcart.Add unEcart
                Next j
            Else
                'Indication que le parcours i ne va pas �tre utilis� et
                'affect� dans le nouvel itin�raire
                unParcours.monIsUtil = False
            End If
        Next i
        
        'Ouverture de la fen�tre d'info des �carts � la moyenne
        'pour confirmation de cr�ation du nouvel itin�raire
        If monNbParSel > 1 Then frmInfoEcart.Show vbModal, Me
        ViderCollection maColEcart
        If monBtnClick = "OK" Or monNbParSel = 1 Then
            'Cas o� on a cliqu� le bouton "OK" de la form frmInfoEcart
            'Cr�ation du nouvel itin�raire avec les valeurs moyennes des
            'parcours ayant des �carts <= � la tol�rance
            'ou qu'il n'y a qu'un parcours s�lectionn�
            '==> On ferme cette fen�tre sans vider maColParcoursMTB
            'et ainsi le frmMain.mnuFileNewByImport_Click appellant
            'cr�era et ouvrira la fen�tre du nouvel itin�raire
            'Si on a cliqu� sur le bouton "Annuler" de la form frmInfoEcart
            '==> On ne fait rien
            Unload Me
        End If
    End If
End Sub

Private Sub btnDeselTout_Click()
    DecocherToutColSelection SpreadParcoursMTB
End Sub

Private Sub btnSelTout_Click()
    CocherToutColSelection SpreadParcoursMTB
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Caption = Format(SpreadParcoursMTB.EditMode)
    If KeyAscii = 27 Then
        'Cas o� l'on tape la touche Echap
        If SpreadParcoursMTB.ActiveCol <> 6 Or SpreadParcoursMTB.EditMode = False Then
            'On fait le traitement du bouton annuler
            'sinon on laisse l'annulation de la frappe en
            'cours dans la saisie du coef d'�talonnage
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
    HelpContextID = HelpID_WinNewByMesure
    
    'Affichage d'un libell� au coin haut gauche
    SpreadParcoursMTB.Row = 0
    SpreadParcoursMTB.Col = 0
    SpreadParcoursMTB.Text = "N�"
    
    'Initialisation du nombre de lignes du spread
    SpreadParcoursMTB.MaxRows = maColParcoursMTB.Count
    'Initialisation de la colonne du coef d'�talonnage du spread
    'avec le s�parateur d�cimal en cours
    TrouverCaract�reD�cimalUtilis�
    InitColSpreadCaract�reD�cimal SpreadParcoursMTB, 6, monCarDeci
    'Affichage du contenu de la collection des parcours du MTB
    'dans le spread de cette fen�tre
    For i = 1 To maColParcoursMTB.Count
        Set unParcours = maColParcoursMTB(i)
        SpreadParcoursMTB.Row = i
        SpreadParcoursMTB.Col = 1
        SpreadParcoursMTB.Text = unParcours.monNom
        SpreadParcoursMTB.Col = 2
        SpreadParcoursMTB.Value = Abs(unParcours.monIsUtil)
        'En effet value = 1 si case coch�e, 0 sinon
        'true = -1 et false = 0 et ICI s�lection = utils�
        SpreadParcoursMTB.Col = 3
        SpreadParcoursMTB.Text = unParcours.monJourSemaine
        SpreadParcoursMTB.Col = 4
        SpreadParcoursMTB.Text = unParcours.maDate
        SpreadParcoursMTB.Col = 5
        SpreadParcoursMTB.Text = unParcours.monHeureDebut
        SpreadParcoursMTB.Col = 6
        SpreadParcoursMTB.Text = unParcours.monCoefEta
        SpreadParcoursMTB.Col = 7
        'R�cup de la distance au dernier top = abs curviligne du dernier top
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
        'Cas de la saisie du coefficient d'�talonnage
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
        'Cas de la saisie du coefficient d'�talonnage
        'R�cup du s�parateur d�cimale en cours
        TrouverCaract�reD�cimalUtilis�
        SpreadParcoursMTB.Row = SpreadParcoursMTB.ActiveRow
        SpreadParcoursMTB.Col = SpreadParcoursMTB.ActiveCol
        If monCarDeci = "." And KeyAscii = 44 Then
            'on remplace la virgule (ascii = 44) par le s�parateur
            'd�cimale en cours pour utiliser le clavier num�rique
            KeyAscii = Asc(monCarDeci)
        ElseIf monCarDeci = "," And KeyAscii = 46 Then
            'on remplace le point (ascii = 46) par le s�parateur
            'd�cimale en cours pour utiliser le clavier num�rique
            KeyAscii = Asc(monCarDeci)
        End If
    End If
End Sub

