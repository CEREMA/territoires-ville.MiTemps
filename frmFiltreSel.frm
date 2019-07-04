VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmFiltreSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtre de sélection des parcours"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frmFiltreSel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnAnnuler 
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
      Left            =   6120
      TabIndex        =   25
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton BtnOK 
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
      Left            =   4200
      TabIndex        =   24
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Frame FrameMeteo 
      Caption         =   "Condition Météo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   5520
      TabIndex        =   13
      Top             =   120
      Width           =   2415
      Begin VB.CheckBox CheckMeteo 
         Caption         =   "Tempête"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CheckBox CheckMeteo 
         Caption         =   "Vent Fort"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   3120
         Width           =   2175
      End
      Begin VB.CheckBox CheckMeteo 
         Caption         =   "Toutes"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox CheckMeteo 
         Caption         =   "Indéfinie"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox CheckMeteo 
         Caption         =   "Beau Temps"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox CheckMeteo 
         Caption         =   "Pluie Forte"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CheckBox CheckMeteo 
         Caption         =   "Pluie Légére"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CheckBox CheckMeteo 
         Caption         =   "Neige"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CheckBox CheckMeteo 
         Caption         =   "Grêle"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CheckBox CheckMeteo 
         Caption         =   "Brouillard"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   2175
      End
   End
   Begin VB.Frame FrameHeure 
      Caption         =   "Heure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1560
      TabIndex        =   11
      Top             =   2160
      Width           =   3855
      Begin FPSpread.vaSpread SpreadHeure 
         Height          =   1260
         Left            =   120
         OleObjectBlob   =   "frmFiltreSel.frx":0442
         TabIndex        =   12
         Top             =   360
         Width           =   3555
      End
   End
   Begin VB.Frame FrameDate 
      Caption         =   "Date"
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
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   3855
      Begin FPSpread.vaSpread SpreadDate 
         Height          =   735
         Left            =   120
         OleObjectBlob   =   "frmFiltreSel.frx":18FA
         TabIndex        =   10
         Top             =   480
         Width           =   3555
      End
      Begin VB.Label Label1 
         Caption         =   "Double-clic possible sur les dates pour ouvrir un calendrier et choisir une date"
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   3585
      End
   End
   Begin VB.Frame FrameJour 
      Caption         =   "Jour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      Begin VB.CheckBox CheckJour 
         Caption         =   "Dimanche"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox CheckJour 
         Caption         =   "Samedi"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox CheckJour 
         Caption         =   "Vendredi"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox CheckJour 
         Caption         =   "Jeudi"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox CheckJour 
         Caption         =   "Mercredi"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox CheckJour 
         Caption         =   "Mardi"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox CheckJour 
         Caption         =   "Lundi"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox CheckJour 
         Caption         =   "Tous"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmFiltreSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAnnuler_Click()
    'Indication que l'on a cliqué sur Annuler pour la fonction appelante
    'ici le bouton filtre de sélection de la fenêtre itinéraire active
    monBtnClick = "Annuler"
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim unNbPar As Integer, unPar As Parcours
    Dim uneSelect As Boolean, unNumJour As Byte
    Dim uneDate1 As Date, uneDate2 As Date, uneDateTmp As Date
    Dim uneH1 As Date, uneH2 As Date
    
    unNbPar = monIti.maColParcours.Count
    For i = 1 To unNbPar
        Set unPar = monIti.maColParcours(i)
        'Test sur le jour
        uneSelect = False
        If CheckJour(0).Value = 1 Then
            'Cas où tous coché pour le jour
            uneSelect = True
        Else
            'Cas des autres jours
            unNumJour = DonnerNumJour(unPar.monJourSemaine)
            If CheckJour(unNumJour).Value = 1 Then uneSelect = True
        End If
        unPar.monIsUtil = uneSelect
        
        If uneSelect Then
            'Test sur la météo
            uneSelect = False
            If CheckMeteo(0).Value = 1 Then
                'Cas où tous coché pour la condition météo
                uneSelect = True
            Else
                'Cas des autres conditions météo
                If CheckMeteo(unPar.maMeteo + 1).Value = 1 Then uneSelect = True
            End If
            unPar.monIsUtil = unPar.monIsUtil And uneSelect
        End If
        
        If uneSelect Then
            'Test sur la date
            uneSelect = False
            SpreadDate.Col = 1
            SpreadDate.Row = 1
            If SpreadDate.Value = 1 Then
                'Cas où tous coché pour la date
                uneSelect = True
            Else
                'Cas des autres plages de date
                For j = 2 To SpreadDate.MaxRows
                    SpreadDate.Row = j
                    If SpreadDate.Value = 1 Then
                        'Cas où la plage est coché
                        'Récup des deux dates et triage croissant
                        SpreadDate.Col = 3
                        uneDate1 = SpreadDate.Text
                        SpreadDate.Col = 5
                        uneDate2 = SpreadDate.Text
                        If uneDate1 > uneDate2 Then
                            uneDateTmp = uneDate1
                            uneDate1 = uneDate2
                            uneDate2 = uneDateTmp
                        End If
                        If unPar.maDate >= uneDate1 And unPar.maDate <= uneDate2 Then
                            uneSelect = True
                            Exit For 'Sortie du for j
                        End If
                    End If
                Next j
            End If
            unPar.monIsUtil = unPar.monIsUtil And uneSelect
        End If
    
        If uneSelect Then
             'Test sur les heures
            uneSelect = False
            SpreadHeure.Col = 1
            SpreadHeure.Row = 1
            If SpreadHeure.Value = 1 Then
                'Cas où tous coché pour l'heure
                uneSelect = True
            Else
                'Cas des autres plages horaires
                For j = 2 To SpreadHeure.MaxRows
                    SpreadHeure.Row = j
                    If SpreadHeure.Value = 1 Then
                        'Cas où la plage est coché
                        'Récup des deux horaires et triage croissant
                        SpreadHeure.Col = 3
                        uneH1 = SpreadHeure.Text
                        SpreadHeure.Col = 5
                        uneH2 = SpreadHeure.Text
                        If uneH1 > uneH2 Then
                            uneDateTmp = uneH1
                            uneH1 = uneH2
                            uneH2 = uneDateTmp
                        End If
                        If unPar.monHeureDebut >= uneH1 And unPar.monHeureDebut <= uneH2 Then
                            uneSelect = True
                            Exit For 'Sortie du for j
                        End If
                    End If
                Next j
            End If
            unPar.monIsUtil = unPar.monIsUtil And uneSelect
        End If
    Next i
    
    If DonnerNbParcoursUtil(monIti) = 1 And monIti.maColParcours(1).monIsUtil Then
        'Cas ou seul reste sélectioné le parcours moyen, donc on ne peut
        'plus le calculer
        MsgBox "Sélection impossible, car il n'aurait plus de parcours sélectionné à part le parcours moyen, donc impossible de calculer les moyennes.", vbExclamation
    Else
        'Indication que l'on a cliqué sur OK pour la fonction appelante
        'ici le bouton filtre de sélection de la fenêtre itinéraire active
        monBtnClick = "OK"
        
        'Initialisation des indicateurs de redessin des onglets de 1 à 6
        'à vrai pour déclencher le dessin lors de leur activation
        IndiquerToutRedessiner monIti
        
        Unload Me
    End If
End Sub

Private Sub CheckJour_Click(Index As Integer)
    If Index = 0 And CheckJour(0).Value = 1 Then
        'Cas du choix de tous on décoche toutes les autres
        For i = 1 To CheckJour.Count - 1
            CheckJour(i).Value = 0
        Next i
    Else
        'Cas du choix d'un jour, on décoche Tous
        If CheckJour(Index).Value = 1 Then CheckJour(0).Value = 0
    End If
End Sub

Private Sub CheckMeteo_Click(Index As Integer)
    If Index = 0 And CheckMeteo(0).Value = 1 Then
        'Cas du choix de tous on décoche toutes les autres
        For i = 1 To CheckMeteo.Count - 1
            CheckMeteo(i).Value = 0
        Next i
    Else
        'Cas du choix d'une condition météo, on décoche Tous
        If CheckMeteo(Index).Value = 1 Then CheckMeteo(0).Value = 0
    End If
End Sub

Private Sub Form_Load()
    CentrerFenetreEcran Me
    'Couleur de fond des cellules lockées
    SpreadDate.LockBackColor = vbInfoBackground
    SpreadHeure.LockBackColor = vbInfoBackground
    
    'Affichage de valeurs par défaut
    CheckJour(0).Value = 1 'Tous coché pour les jours
    'Tous coché pour la date
    SpreadDate.Row = 1
    SpreadDate.Col = 1
    SpreadDate.Value = 1
    'Deux Plages de dates par défaut
    'Plage 1 = entre hier et aujourd'hui
    'Plage 2 = entre avant-hier et hier
    SpreadDate.Row = 2
    SpreadDate.Col = 3
    SpreadDate.Text = Date - 1
    SpreadDate.Col = 5
    SpreadDate.Text = Date
    
    SpreadDate.Row = 3
    SpreadDate.Col = 3
    SpreadDate.Text = Date - 2
    SpreadDate.Col = 5
    SpreadDate.Text = Date - 1
    
    'Tous coché pour l'heure
    SpreadHeure.Row = 1
    SpreadHeure.Col = 1
    SpreadHeure.Value = 1
    'Quatre plages d'heures par défaut, les 3 périodes
    'de pointes 7-9, 12-14, 17-19 plus une autre 20-22
    SpreadHeure.Row = 2
    SpreadHeure.Col = 3
    SpreadHeure.Text = "07:00:00"
    SpreadHeure.Col = 5
    SpreadHeure.Text = "09:00:00"
    
    SpreadHeure.Row = 3
    SpreadHeure.Col = 3
    SpreadHeure.Text = "12:00:00"
    SpreadHeure.Col = 5
    SpreadHeure.Text = "14:00:00"
    
    SpreadHeure.Row = 4
    SpreadHeure.Col = 3
    SpreadHeure.Text = "17:00:00"
    SpreadHeure.Col = 5
    SpreadHeure.Text = "19:00:00"
    
    SpreadHeure.Row = 5
    SpreadHeure.Col = 3
    SpreadHeure.Text = "20:00:00"
    SpreadHeure.Col = 5
    SpreadHeure.Text = "22:00:00"
    
    'Affichage des libellés des conditions météo
    CheckMeteo(0).Value = 1 'Tous coché pour la météo
    For i = 1 To 9
        CheckMeteo(i).Caption = Mid(monIti.maColMeteo(i), 5)
    Next i
End Sub



Private Sub SpreadDate_Advance(ByVal AdvanceNext As Boolean)
    If AdvanceNext Then
        'On passe le focus au control suivant de la feuille
        SpreadHeure.SetFocus
    Else
        'On passe le focus au control précédent de la feuille
        CheckJour(CheckJour.Count - 1).SetFocus
    End If
End Sub

Private Sub SpreadDate_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    SpreadDate.Row = Row
    SpreadDate.Col = Col
    If Row = 1 And SpreadDate.Value = 1 Then
        'Cas du choix de tous on décoche toutes les autres
        For i = 2 To SpreadDate.MaxRows
            SpreadDate.Row = i
            SpreadDate.Value = 0
        Next i
    Else
        'Cas du choix d'une plage de date, on décoche Tous
        If SpreadDate.Value = 1 Then
            SpreadDate.Row = 1
            SpreadDate.Value = 0
        End If
    End If
End Sub

Private Sub SpreadHeure_Advance(ByVal AdvanceNext As Boolean)
    If AdvanceNext = False Then
        'On passe le focus au control suivant de la feuille
        SpreadDate.SetFocus
    Else
        'On passe le focus au control précédent de la feuille
        CheckMeteo(0).SetFocus
    End If
End Sub

Private Sub SpreadHeure_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    SpreadHeure.Row = Row
    SpreadHeure.Col = Col
    If Row = 1 And SpreadHeure.Value = 1 Then
        'Cas du choix de tous on décoche toutes les autres
        For i = 2 To SpreadHeure.MaxRows
            SpreadHeure.Row = i
            SpreadHeure.Value = 0
        Next i
    Else
        'Cas du choix d'une plage de date, on décoche Tous
        If SpreadHeure.Value = 1 Then
            SpreadHeure.Row = 1
            SpreadHeure.Value = 0
        End If
    End If
End Sub

Private Function DonnerNumJour(unJour As String) As Byte
    'Retourne le numéro du jour dans la semaine
    'de 1 à 7  = de lundi à dimanche
    unJour = LCase(unJour)
    If unJour = "lundi" Then
        DonnerNumJour = 1
    ElseIf unJour = "lundi" Then
        DonnerNumJour = 1
    ElseIf unJour = "mardi" Then
        DonnerNumJour = 2
    ElseIf unJour = "mercredi" Then
        DonnerNumJour = 3
    ElseIf unJour = "jeudi" Then
        DonnerNumJour = 4
    ElseIf unJour = "vendredi" Then
        DonnerNumJour = 5
    ElseIf unJour = "samedi" Then
        DonnerNumJour = 6
    ElseIf unJour = "dimanche" Then
        DonnerNumJour = 7
    Else
        DonnerNumJour = 0
    End If
End Function
