VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmInfoEcart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informations sur les �carts des parcours s�lectionn�s"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   Icon            =   "frmInfoEcart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread SpreadInfoEcart 
      Height          =   3255
      Left            =   120
      OleObjectBlob   =   "frmInfoEcart.frx":0442
      TabIndex        =   1
      Top             =   480
      Width           =   9735
   End
   Begin VB.CommandButton btnCancel 
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
      Left            =   5040
      TabIndex        =   3
      Top             =   5040
      Width           =   4815
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
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
      Width           =   4815
   End
   Begin VB.Label LabelInfoUser 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      TabIndex        =   4
      Top             =   4440
      Width           =   585
   End
   Begin VB.Label LabelInfoTol 
      AutoSize        =   -1  'True
      Caption         =   $"frmInfoEcart.frx":16FC
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "frmInfoEcart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
    ViderCollection maColValMoy
    monBtnClick = "Annuler"
    Unload Me
End Sub

Private Sub btnOK_Click()
    monBtnClick = "OK"
    Unload Me
End Sub

Private Sub Form_Load()
    Dim unPar As Parcours, unIndRep As Integer
    Dim unIndEcart As Integer, unEcart As Boolean
    Dim unNbParAvecEcart As Integer
    
    CentrerFenetreEcran Me
    unNbParAvecEcart = 0
    LabelInfoTol.Caption = "En rouge sont signal�s les �carts (= |Valeur - Moyenne| / Moyenne)  > Tol�rance �cart en longueur (= " + Format(mesOptions.maTolLong) + " % cf menu Affichage / Options)"
    
    'Contexte d'aide
    HelpContextID = HelpID_WinNewByMesure
    
    'Nb lignes = nombre de parcours s�lectionn�s dans la form
    'frmNewParMTB + le parcours moyen
    SpreadInfoEcart.MaxRows = frmNewParMTB.monNbParSel + 1
    'Nb col = 2 + nb rep�re * 2 (= nb total de valeurs moyennes)
    'car � la fin on a la distance parcourue et la dur�e de mesure
    SpreadInfoEcart.MaxCols = maColValMoy.Count * 2 + 2
    'Mise de la couleur de fond = couleur info bulle
    'car non modifiable
    SpreadInfoEcart.Col = -1
    SpreadInfoEcart.Row = -1
    SpreadInfoEcart.BackColor = vbInfoBackground
    
    'R�cup du nombre de rep�res
    unNbRep = (SpreadInfoEcart.MaxCols - 2) / 2
    
    'Remplissage du coin haut gauche
    SpreadInfoEcart.Row = 0
    SpreadInfoEcart.Col = 0
    SpreadInfoEcart.Text = "N�"
    
    'Remplissage de la 1�re ligne avec le parcours moyen
    SpreadInfoEcart.Row = 1
    SpreadInfoEcart.Col = 0
    SpreadInfoEcart.Text = "0"
    SpreadInfoEcart.Col = 1
    SpreadInfoEcart.Text = "Parcours moyen"
    For j = 1 To unNbRep
        'Affichage des ent�tes de colonnes et des valeurs moyennes
        SpreadInfoEcart.Col = j * 2 + 1
        SpreadInfoEcart.Row = 0
        SpreadInfoEcart.Text = "Abs curv Rep�re " + Format(j) + " (m)"
        
        SpreadInfoEcart.Row = 1
        SpreadInfoEcart.TypeHAlign = 2 'texte centr�
        SpreadInfoEcart.ColWidth(j * 2 + 1) = 1030
        SpreadInfoEcart.Text = Format(maColValMoy(j), "fixed")
        
        SpreadInfoEcart.Row = 0
        SpreadInfoEcart.Col = j * 2 + 2
        SpreadInfoEcart.ColWidth(j * 2 + 2) = 1030
        SpreadInfoEcart.Text = "Ecart avec la moyenne (%)"
        SpreadInfoEcart.Row = 1
        SpreadInfoEcart.Text = ""
    Next j
    
    'Remplissage des colonnes des parcours choisis � partir de la ligne 2
    k = 0
    For i = 1 To maColParcoursMTB.Count
        Set unPar = maColParcoursMTB(i)
        unEcart = False
        If unPar.monIsUtil Then
            'Remplissage de la ligne k
            k = k + 1
            'Affichage des abs curv des rep�res et des �carts
            For j = 1 To unNbRep
                'Affichage du num�ro de ligne du parcours du spread
                'de la form frmNewParMTB
                SpreadInfoEcart.Row = k + 1
                SpreadInfoEcart.Col = 0
                SpreadInfoEcart.Text = Format(i)
                
                'Affichage des abs curv des rep�res
                unIndEcart = maColValMoy.Count * (k - 1) + j
                SpreadInfoEcart.Row = k + 1
                SpreadInfoEcart.Col = j * 2 + 1
                SpreadInfoEcart.TypeHAlign = 2 'texte centr�
                SpreadInfoEcart.ColWidth(j * 2 + 1) = 1030
                SpreadInfoEcart.Text = Format(CLng(unPar.monTabAbsRep(j) * unPar.monCoefEta) / 10)
                If frmNewParMTB.maColEcart(unIndEcart) > mesOptions.maTolLong Then
                    SpreadInfoEcart.ForeColor = LabelInfoTol.ForeColor
                    unEcart = True
                End If
                
                'Affichage des �carts des rep�res
                SpreadInfoEcart.Row = k + 1
                SpreadInfoEcart.Col = j * 2 + 2
                SpreadInfoEcart.TypeHAlign = 2 'texte centr�
                SpreadInfoEcart.ColWidth(j * 2 + 2) = 1030
                SpreadInfoEcart.Text = Format(frmNewParMTB.maColEcart(unIndEcart), "fixed")
            Next j
            'Indication si un �cart avec la moyenne a �t� trouv�
            SpreadInfoEcart.Row = k + 1
            SpreadInfoEcart.Col = 2
            SpreadInfoEcart.Value = Abs(unEcart) '1 = coch�e, 0 sinon
            'Nom mis en rouge si un �cart avec la moyenne a �t� trouv�
            SpreadInfoEcart.Col = 1
            If unEcart Then
                SpreadInfoEcart.ForeColor = LabelInfoTol.ForeColor
                'Incr�mentation du nombre de parcours ayant �cart > tol�rance
                unNbParAvecEcart = unNbParAvecEcart + 1
            End If
            SpreadInfoEcart.Text = unPar.monNom
        End If
    Next i
    
    If unNbParAvecEcart > 0 Then
        'Au moins un parcours s�lectionn� a un �cart > Tol�rance
        '===> Cr�ation impossible
        LabelInfoUser.Caption = "Impossible de cr�er l'itin�raire de r�f�rence car certains parcours ont un �cart > � la tol�rance."
        BtnOK.Visible = False
        btnCancel.Width = ScaleWidth - 240
        btnCancel.Left = 120
    Else
        'Cas o� tous les parcours s�lectionn�s ont un �cart <= tol�rance
        'On calculera le parcours moyen
        LabelInfoUser.Caption = "Voulez-vous cr�er l'itin�raire de r�f�rence en prenant les valeurs " + Chr(13) + "du parcours moyen (ligne N� 0) et en y important les parcours ci-dessus ?"
        BtnOK.Caption = "Oui"
        btnCancel.Caption = "Non"
    End If
End Sub


