VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options de MiTemps"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin VB.Frame FrameMeteo 
      Caption         =   "Conditions météo du boitier "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   7920
      TabIndex        =   38
      Top             =   120
      Width           =   2800
      Begin FPSpread.vaSpread SpreadMeteo 
         Height          =   2415
         Left            =   120
         OleObjectBlob   =   "frmOptions.frx":0442
         TabIndex        =   39
         Top             =   360
         Width           =   2600
      End
   End
   Begin VB.Frame FrameClasseV 
      Caption         =   "Couleurs des différentes classes de vitesses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   7695
      Begin VB.PictureBox PictureClasV4 
         BackColor       =   &H0000FF00&
         Height          =   315
         Left            =   6960
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   37
         Top             =   300
         Width           =   615
      End
      Begin VB.ComboBox ComboVal5 
         Height          =   315
         ItemData        =   "frmOptions.frx":0A15
         Left            =   5715
         List            =   "frmOptions.frx":0A17
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   750
         Width           =   645
      End
      Begin VB.ComboBox ComboVal4 
         Height          =   315
         ItemData        =   "frmOptions.frx":0A19
         Left            =   5715
         List            =   "frmOptions.frx":0A1B
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   300
         Width           =   645
      End
      Begin VB.ComboBox ComboVal3 
         Height          =   315
         ItemData        =   "frmOptions.frx":0A1D
         Left            =   1185
         List            =   "frmOptions.frx":0A1F
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1200
         Width           =   645
      End
      Begin VB.ComboBox ComboVal2 
         Height          =   315
         ItemData        =   "frmOptions.frx":0A21
         Left            =   1185
         List            =   "frmOptions.frx":0A23
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   750
         Width           =   645
      End
      Begin VB.ComboBox ComboVal1 
         Height          =   315
         ItemData        =   "frmOptions.frx":0A25
         Left            =   1200
         List            =   "frmOptions.frx":0A27
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   300
         Width           =   645
      End
      Begin VB.PictureBox PictureClasV6 
         BackColor       =   &H00FF0000&
         Height          =   315
         Left            =   6960
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   36
         Top             =   1200
         Width           =   615
      End
      Begin VB.PictureBox PictureClasV5 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   6960
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   34
         Top             =   750
         Width           =   615
      End
      Begin VB.PictureBox PictureClasV3 
         BackColor       =   &H0000FFFF&
         Height          =   315
         Left            =   2430
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   27
         Top             =   1200
         Width           =   615
      End
      Begin VB.PictureBox PictureClasV2 
         BackColor       =   &H000080FF&
         Height          =   315
         Left            =   2430
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   23
         Top             =   750
         Width           =   615
      End
      Begin VB.PictureBox PictureClasV1 
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   2430
         ScaleHeight     =   255
         ScaleWidth      =   555
         TabIndex        =   19
         Top             =   300
         Width           =   615
      End
      Begin VB.Label LabelClasV6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plus de 000 km/h :"
         Height          =   195
         Left            =   5520
         TabIndex        =   35
         Top             =   1260
         Width           =   1350
      End
      Begin VB.Label Labelkm5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "km/h : "
         Height          =   195
         Left            =   6420
         TabIndex        =   33
         Top             =   810
         Width           =   510
      End
      Begin VB.Label LabelClasV5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plus de 000 à  "
         Height          =   195
         Left            =   4650
         TabIndex        =   31
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label Labelkm4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "km/h : "
         Height          =   195
         Left            =   6420
         TabIndex        =   30
         Top             =   360
         Width           =   510
      End
      Begin VB.Label LabelClasV4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plus de 000 à  "
         Height          =   195
         Left            =   4650
         TabIndex        =   28
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Labelkm3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "km/h : "
         Height          =   195
         Left            =   1890
         TabIndex        =   26
         Top             =   1260
         Width           =   510
      End
      Begin VB.Label LabelClasV3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plus de 000 à  "
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Labelkm2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "km/h : "
         Height          =   195
         Left            =   1890
         TabIndex        =   22
         Top             =   810
         Width           =   510
      End
      Begin VB.Label Labelkm1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "km/h : "
         Height          =   195
         Left            =   1890
         TabIndex        =   18
         Top             =   360
         Width           =   510
      End
      Begin VB.Label LabelClasV2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plus de 000 à  "
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label LabelClasV1 
         AutoSize        =   -1  'True
         Caption         =   "De 0 à  "
         Height          =   195
         Left            =   600
         TabIndex        =   16
         Top             =   360
         Width           =   570
      End
   End
   Begin VB.Frame FrameTolLong 
      Caption         =   "Tolérance sur les longueurs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7695
      Begin VB.ComboBox ComboEcartDTop 
         Height          =   315
         ItemData        =   "frmOptions.frx":0A29
         Left            =   6360
         List            =   "frmOptions.frx":0A4B
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox ComboTolLg 
         Height          =   315
         ItemData        =   "frmOptions.frx":0A6E
         Left            =   6360
         List            =   "frmOptions.frx":0A90
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "mètres"
         Height          =   195
         Left            =   7080
         TabIndex        =   14
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   7080
         TabIndex        =   11
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ecart maximun avec un repère pour la détection d'un double top :"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   4620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pourcentage d'écart entre les longueurs d'un même parcours et sa longueur moyenne : "
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   6165
      End
   End
   Begin VB.CommandButton cmdOK 
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
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   3240
      Width           =   1935
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
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Tag             =   "Annuler"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Exemple 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   7
         Tag             =   "Exemple 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Exemple 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   6
         Tag             =   "Exemple 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Exemple 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   4
         Tag             =   "Exemple 2"
         Top             =   305
         Width           =   2033
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
    FermerFenetre Me
End Sub


Private Sub cmdOK_Click()
    Dim unMsgErreur As String
    Dim uneErreurClasV As Boolean
    
    'Vérification de la contiguité des plages de valeurs des classes
    'de vitesse
    uneErreurClasV = False
    If CInt(ComboVal1.Text) >= CInt(ComboVal2.Text) Then
        uneErreurClasV = True
    ElseIf CInt(ComboVal2.Text) >= CInt(ComboVal3.Text) Then
        uneErreurClasV = True
    ElseIf CInt(ComboVal3.Text) >= CInt(ComboVal4.Text) Then
        uneErreurClasV = True
    ElseIf CInt(ComboVal4.Text) >= CInt(ComboVal5.Text) Then
        uneErreurClasV = True
    End If
    
    If uneErreurClasV = False Then
        'Pas d'erreur dans la partition des classes de vitesses
        'donc on valide les modifs
        
        'Modification des options
        With mesOptions
            .maTolLong = CByte(ComboTolLg.Text)
            .monEcartMax = CByte(ComboEcartDTop.Text)
            .maCouleurClasV1 = PictureClasV1.BackColor
            .maValClasV1 = Format(ComboVal1.Text)
            .maCouleurClasV2 = PictureClasV2.BackColor
            .maValClasV2 = Format(ComboVal2.Text)
            .maCouleurClasV3 = PictureClasV3.BackColor
            .maValClasV3 = Format(ComboVal3.Text)
            .maCouleurClasV4 = PictureClasV4.BackColor
            .maValClasV4 = Format(ComboVal4.Text)
            .maCouleurClasV5 = PictureClasV5.BackColor
            .maValClasV5 = Format(ComboVal5.Text)
            .maCouleurClasV6 = PictureClasV6.BackColor
        
            uneStrTmp = ""
            SpreadMeteo.Col = 1
            For i = 0 To 7
                SpreadMeteo.Row = i + 1
                uneStrTmp = uneStrTmp + SpreadMeteo.Text + ","
            Next i
            SpreadMeteo.Row = i + 1
            .mesLibMeteo = uneStrTmp + SpreadMeteo.Text
        End With
        
        'Stockage des options dans la base de registre
        StockerOptions
        
        FermerFenetre Me
    Else
        'Cas de détection d'erreur dans la partition des classes
        'de vitesses
        unMsgErreur = "Les bornes de fin des classes de vitesses doivent vérifiées l'ordonnancement suivant :" + Chr(13) + Chr(13)
        unMsgErreur = unMsgErreur + "   0 < Borne Fin Classe 1 < Borne Fin Classe 2 < Borne Fin Classe 3 < Borne Fin Classe 4 < Borne Fin Classe 5" + Chr(13)
        unMsgErreur = unMsgErreur + Chr(13) + "Or votre ordonnancement 0 < " + ComboVal1.Text + " < " + ComboVal2.Text + " < " + ComboVal3.Text + " < " + ComboVal4.Text + " < " + ComboVal5.Text
        unMsgErreur = unMsgErreur + " ne vérifie pas cet ordre."
        MsgBox unMsgErreur, vbCritical, App.Title + " : Mauvais partitionnement des classes de vitesses"
    End If
End Sub


Private Sub ComboVal1_Click()
    ModifierLabelClass ComboVal1, ComboVal2, LabelClasV2
End Sub

Private Sub ComboVal2_Click()
    ModifierLabelClass ComboVal2, ComboVal3, LabelClasV3
End Sub

Private Sub ComboVal3_Click()
    ModifierLabelClass ComboVal3, ComboVal4, LabelClasV4
End Sub

Private Sub ComboVal4_Click()
    ModifierLabelClass ComboVal4, ComboVal5, LabelClasV5
End Sub

Private Sub ComboVal5_Click()
    ModifierLabelClass ComboVal5, ComboVal5, LabelClasV6
End Sub

Private Sub Form_Load()
    'Bouton OK actif uniquement si aucune fenêtre fille d'ouverte
    'donc 2 fenêtres ouvertes, la MDI mère et les options
    cmdOK.Enabled = (Forms.Count = 2)
    
    'Contexte d'aide
    HelpContextID = HelpID_WinOptions
    
    'Remplissage des combobox des classes de vitesses
    For i = 1 To 15
        ComboVal1.AddItem Format(i * 10)
        ComboVal2.AddItem Format(i * 10)
        ComboVal3.AddItem Format(i * 10)
        ComboVal4.AddItem Format(i * 10)
        ComboVal5.AddItem Format(i * 10)
    Next i
    
    'Affectations des valeurs stockées dans les options
    ComboTolLg.Text = Format(mesOptions.maTolLong)
    ComboEcartDTop.Text = Format(mesOptions.monEcartMax)
    
    PictureClasV1.BackColor = Format(mesOptions.maCouleurClasV1)
    ComboVal1.Text = Format(mesOptions.maValClasV1)
    PictureClasV2.BackColor = Format(mesOptions.maCouleurClasV2)
    ComboVal2.Text = Format(mesOptions.maValClasV2)
    PictureClasV3.BackColor = Format(mesOptions.maCouleurClasV3)
    ComboVal3.Text = Format(mesOptions.maValClasV3)
    PictureClasV4.BackColor = Format(mesOptions.maCouleurClasV4)
    ComboVal4.Text = Format(mesOptions.maValClasV4)
    PictureClasV5.BackColor = Format(mesOptions.maCouleurClasV5)
    ComboVal5.Text = Format(mesOptions.maValClasV5)
    PictureClasV6.BackColor = Format(mesOptions.maCouleurClasV6)
    
    unePos0 = 1
    SpreadMeteo.Col = 1
    For i = 0 To 7
        'Recherche de la virgule séparant les libellés météo
        unePos = InStr(unePos0, mesOptions.mesLibMeteo, ",")
        SpreadMeteo.Row = i + 1
        SpreadMeteo.Text = Mid(mesOptions.mesLibMeteo, unePos0, unePos - unePos0)
        unePos0 = unePos + 1
    Next i
    SpreadMeteo.Row = i + 1
    SpreadMeteo.Text = Mid(mesOptions.mesLibMeteo, unePos0)
End Sub







Private Sub PictureClasV1_Click()
    ChoisirCouleur PictureClasV1
End Sub

Private Sub PictureClasV2_Click()
    ChoisirCouleur PictureClasV2
End Sub

Private Sub PictureClasV3_Click()
    ChoisirCouleur PictureClasV3
End Sub

Private Sub PictureClasV4_Click()
    ChoisirCouleur PictureClasV4
End Sub

Private Sub PictureClasV5_Click()
    ChoisirCouleur PictureClasV5
End Sub

Private Sub PictureClasV6_Click()
    ChoisirCouleur PictureClasV6
End Sub

Private Sub ModifierLabelClass(uneComboText As ComboBox, uneComboLeft As ComboBox, unLabelClasV As Label)
    'Modification du text et du placement du label passé en paramètres
    'Le texte est modifié par le text de la combobox uneComboText
    'et le placement est à coté à droite de la combobox uneComboLeft
    If unLabelClasV Is LabelClasV6 Then
        unLabelClasV.Caption = "Plus de " + uneComboText.Text + " km/h : "
        unLabelClasV.Left = PictureClasV6.Left - unLabelClasV.Width - 60
    Else
        unLabelClasV.Caption = "Plus de " + uneComboText.Text + " à "
        unLabelClasV.Left = uneComboLeft.Left - unLabelClasV.Width - 60
    End If
End Sub

Private Sub SpreadMeteo_GotFocus()
    cmdOK.Default = False
    cmdCancel.Cancel = False
End Sub

Private Sub SpreadMeteo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And SpreadMeteo.EditMode = False Then
        FermerFenetre Me
    End If
    
    If SpreadMeteo.ActiveRow = SpreadMeteo.MaxRows Then
        If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
            'On remet le OK en bouton par défaut avec le focus
            'et le cancel par défaut
            cmdOK.SetFocus
            cmdOK.Default = True
            cmdCancel.Cancel = True
        End If
    End If
End Sub

Private Sub SpreadMeteo_LostFocus()
    'On rend actif la ligne 1
    SpreadMeteo.Row = 0
    SpreadMeteo.Col = 1
    SpreadMeteo.Action = 0 'SS_CELL_ACTIVE
End Sub
