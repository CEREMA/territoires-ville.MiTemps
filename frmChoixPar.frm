VERSION 5.00
Begin VB.Form frmChoixPar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parcours trouv�s"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "frmChoixPar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnAnnuler 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1960
      Width           =   975
   End
   Begin VB.CommandButton btnChoisir 
      Caption         =   "Choisir"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1960
      Width           =   975
   End
   Begin VB.ListBox ListParTrouv 
      Height          =   1815
      ItemData        =   "frmChoixPar.frx":0442
      Left            =   60
      List            =   "frmChoixPar.frx":044F
      TabIndex        =   0
      Top             =   60
      Width           =   5475
   End
End
Attribute VB_Name = "frmChoixPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAnnuler_Click()
    If Tag = "" Then
        'Remise � vide du tag de la fen�tre import de campagne de mesures
        frmImportMTB.Tag = ""
    ElseIf Tag = "DT" Or Tag = "DV" Then
        'Remise � vide du tag de la picture active (courbe DT ou DT) de la
        'fen�tre itin�raire active
        maPicBox.Tag = ""
    Else
        MsgBox MsgErreurProg + MsgErreurTypeTagInconnu + MsgIn + "frmChoixPar:btnAnnuler_Click", vbCritical
        Exit Sub
    End If
    
    monBtnClick = "Annuler" + Format(ListParTrouv.ListIndex)
    Unload Me
End Sub

Private Sub btnChoisir_Click()
    If ListParTrouv.ListIndex = -1 Then
        MsgBox "Vous devez choisir un parcours parmi ceux affich�s", vbInformation
    Else
        'Stockage dans le tag de la fen�tre d'import campagne de mesure
        'du parcours s�lectionn� dans la listbox ListParTrouv issu du fichier MTB
        If Tag = "" Then
            'Stockage dans le tag de la fen�tre d'import campagne de mesure
            'du parcours s�lectionn� dans la listbox ListParTrouv issu du
            'fichier MTB
            frmImportMTB.Tag = Format(ListParTrouv.ItemData(ListParTrouv.ListIndex))
        ElseIf Tag = "DT" Or Tag = "DV" Then
            'Stockage dans le tag de la picture active (courbe DT ou DT)de la
            'fen�tre itin�raire active du parcours s�lectionn� dans la listbox
            'ListParTrouv issu du fichier MTB
            maPicBox.Tag = Format(ListParTrouv.ItemData(ListParTrouv.ListIndex))
        Else
            MsgBox MsgErreurProg + MsgErreurTypeTagInconnu + MsgIn + "frmChoixPar:btnChoisir_Click", vbCritical
            Exit Sub
        End If
        Unload Me
    End If
End Sub


Private Sub Form_Activate()
    monBtnClick = ""
    'Remise � vide du tag de la fen�tre import de campagne de mesures
    'si frmChoixPar appel�e par click dans frmImportMTB,
    'sinon Tag = "DV" si frmChoixPar appel�e par click dans le dessin
    'd'une courbe DV dans une fen�tre Itin�raire ou "DT" dans une courbe DT
    If Tag = "" Then frmImportMTB.Tag = ""
    'Calcul de la largeur maxi des textes dans la listbox
    uneWidthMax = 0
    For i = 0 To ListParTrouv.ListCount - 1
        If Me.TextWidth(ListParTrouv.List(i)) > uneWidthMax Then
            uneWidthMax = Me.TextWidth(ListParTrouv.List(i))
        End If
    Next i
    'On augmente la largeur maxi pour ne rien masquer
    'si un scroller vertical apparait (+ 300 twips)
    uneWidthMax = uneWidthMax + 420
    
    'Retaillage et placement des controles de la fen�tre
    Me.Width = uneWidthMax + ListParTrouv.Left + 180
    ListParTrouv.Width = uneWidthMax
    uneMarge = Me.ScaleWidth - btnChoisir.Width - btnAnnuler.Width - 120
    '120 = espacement entre les deux boutons Choisir et Annuler
    btnChoisir.Left = uneMarge / 2
    btnAnnuler.Left = btnChoisir.Left + btnChoisir.Width + 120
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim unIndParChoisi As Integer
    
    'Si un parcours a �t� choisi et que la fen�tre de choix a �t� ferm� par
    'click dans le menu syst�me ou la croix en haut � droite,
    'on redessine tous en trait fin en indiquant que l'indice du parcours
    'choisi vaut 0 dans frmImport, on indique qu'aucun parcours n' a �t� choisi
    If ListParTrouv.ListIndex > -1 And UnloadMode = vbFormControlMenu Then
        Screen.MousePointer = vbHourglass
        'Mise � z�ro de l'indice dans la liste des parcours issus du fichier MTB
        'du parcours s�lectionn� dans la listbox ListParTrouv, aucun parcours choisi
        unIndParChoisi = 0
        If Tag = "DT" Or Tag = "DV" Then
            'On indique qu'aucun parcours n'a �t� choisi � la fen�tre
            'itin�raire active
            maPicBox.Tag = ""
        Else
            'Affichage en trait �pais du parcours choisi dans la zone de dessin de la
            'fen�tre import d'une campagne de mesure, or aucun parcours choisi
            'donc ils seront tous dessin�s en trait fin
            frmImportMTB.MontrerParcoursChoisi unIndParChoisi
        End If
        'Remise � vide du tag de la list box
        ListParTrouv.Tag = ""
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub ListParTrouv_Click()
    Dim unIndParChoisi As Integer
    
    'Si on s�lectionne l'item d�j� s�lectionn� on ne fait rien
    If ListParTrouv.Tag = Format(ListParTrouv.ListIndex) Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    'R�cup�ration de l'indice dans la liste des parcours issus du fichier MTB
    'du parcours s�lectionn� dans la listbox ListParTrouv
    unIndParChoisi = ListParTrouv.ItemData(ListParTrouv.ListIndex)
    
    If Tag = "DT" Then
        'On indique quel indice de parcours a �t� choisi dans la picture box
        'PicBoxDT de la fen�tre itin�raire active
        maPicBox.Tag = Format(unIndParChoisi)
        monIti.monIndParcoursSelectDT = unIndParChoisi
        DessinerCourbe monIti, maPicBox, OngletCbeDT
    ElseIf Tag = "DV" Then
        'On indique quel indice de parcours a �t� choisi dans la picture box
        'PicBoxDV de la fen�tre itin�raire active
        maPicBox.Tag = Format(unIndParChoisi)
        monIti.monIndParcoursSelectDV = unIndParChoisi
        DessinerCourbe monIti, maPicBox, OngletCbeDV
    Else
        'Affichage en trait �pais du parcours choisi dans la zone de dessin de la
        'fen�tre import d'une campagne de mesure
        frmImportMTB.MontrerParcoursChoisi unIndParChoisi

    End If
    
    'Stockage de l'item s�lectionn�
    ListParTrouv.Tag = Format(ListParTrouv.ListIndex)
    Screen.MousePointer = vbDefault
End Sub
