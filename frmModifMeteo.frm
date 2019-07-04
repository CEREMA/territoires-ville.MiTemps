VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmModifMeteo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modification des libellés des conditions Météo"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmModifMeteo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread SpreadMeteo 
      Height          =   2415
      Left            =   960
      OleObjectBlob   =   "frmModifMeteo.frx":0442
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton btnAnnuler 
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
      Left            =   2400
      TabIndex        =   3
      Top             =   3720
      Width           =   2175
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
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   $"frmModifMeteo.frx":095B
      Height          =   795
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4545
   End
End
Attribute VB_Name = "frmModifMeteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAnnuler_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim uneListeMeteo As String
    
    SpreadMeteo.Col = 1
    For i = 0 To 8
        SpreadMeteo.Row = i + 1
        If Trim(SpreadMeteo.Text) = "" Then
            MsgBox "Aucun libellé de conditions météo ne peut être vide.", vbCritical
            Exit Sub
        End If
        'Modif de la collection des libellés météo de la
        'fenêtre itinéraire active
        monIti.maColMeteo.Add Format(i) + " - " + SpreadMeteo.Text, , i + 1
        monIti.maColMeteo.Remove i + 2
        'Création de la chaine contenant tous les libellés
        'séparées par le caractére de code ASCII 9 pour
        'intégration dans une cellule combobox de spread
        uneListeMeteo = uneListeMeteo + Format(i) + " - " + SpreadMeteo.Text
        If i < 8 Then uneListeMeteo = uneListeMeteo + Chr(9)
    Next i
    
    'Modification du contenu des combox de chaque parcours
    '= chaque ligne du spread parcours
    monIti.SpreadParcours.Col = 6
    For i = 1 To monIti.SpreadParcours.MaxRows
        monIti.SpreadParcours.Row = i
        unCurSel = monIti.SpreadParcours.TypeComboBoxCurSel
        monIti.SpreadParcours.TypeComboBoxList = uneListeMeteo
        monIti.SpreadParcours.TypeComboBoxCurSel = unCurSel
    Next i
    
    'Indication de modif
    monIti.maModif = True
    
    Unload Me
End Sub

Private Sub Form_Load()
    'Chargement avec les libellés de la fenêtre
    'itinéraire active
    CentrerFenetreEcran Me
    SpreadMeteo.Col = 1
    For i = 0 To 8
        SpreadMeteo.Row = i + 1
        SpreadMeteo.Text = Mid(monIti.maColMeteo(i + 1), 5)
    Next i
End Sub

Private Sub SpreadMeteo_GotFocus()
    btnOK.Default = False
    btnAnnuler.Cancel = False
End Sub

Private Sub SpreadMeteo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And SpreadMeteo.EditMode = False Then
        FermerFenetre Me
    End If
    
    If SpreadMeteo.ActiveRow = SpreadMeteo.MaxRows Then
        If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
            'On remet le OK en bouton par défaut avec le focus
            'et le cancel par défaut
            btnOK.SetFocus
            btnOK.Default = True
            btnAnnuler.Cancel = True
        End If
    End If

End Sub

Private Sub SpreadMeteo_LostFocus()
    'On rend actif la ligne 1
    SpreadMeteo.Row = 0
    SpreadMeteo.Col = 1
    SpreadMeteo.Action = 0 'SS_CELL_ACTIVE
End Sub
