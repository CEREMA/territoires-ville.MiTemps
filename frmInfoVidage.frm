VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form frmInfoVidage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Information sur les parcours récupérés du boitier"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   Icon            =   "frmInfoVidage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread SpreadInfoVidage 
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "frmInfoVidage.frx":0442
      TabIndex        =   1
      Top             =   120
      Width           =   10575
   End
   Begin VB.CommandButton btnFermer 
      Cancel          =   -1  'True
      Caption         =   "Fermer"
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
      TabIndex        =   0
      Top             =   5040
      Width           =   10575
   End
End
Attribute VB_Name = "frmInfoVidage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnFermer_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Tag = "A fermer" Then Unload Me
End Sub

Private Sub Form_Load()
    Dim unParcours As Parcours
    
    Me.MousePointer = vbHourglass
    SpreadInfoVidage.BackColor = vbInfoBackground
    CentrerFenetreEcran Me
    'Contexte d'aide
    HelpContextID = HelpID_WinViderBoitier
    'On vide les parcours issus du mtb
    ViderColParcours maColParcoursMTB
    If LireFichierMTB(frmViderBoitier.FichierOuvrir) Then
        'Cas où la lecture du MTB s'est bien passée
        'Affichage du contenu de la collection des parcours du MTB
        'dans le spread de cette fenêtre
        SpreadInfoVidage.MaxRows = maColParcoursMTB.Count
        For i = 1 To maColParcoursMTB.Count
            Set unParcours = maColParcoursMTB(i)
            SpreadInfoVidage.Row = i
            SpreadInfoVidage.Col = 1
            SpreadInfoVidage.Text = unParcours.monNom
            SpreadInfoVidage.Col = 2
            SpreadInfoVidage.Text = unParcours.monJourSemaine
            SpreadInfoVidage.Col = 3
            SpreadInfoVidage.Text = unParcours.maDate
            SpreadInfoVidage.Col = 4
            SpreadInfoVidage.Text = unParcours.monHeureDebut
            
            SpreadInfoVidage.Col = 5
            'Formattage en 00h 00mn 00s de la durée
            uneStringDuree = FormatterTempsEnHMNS(unParcours.maDuree)
            SpreadInfoVidage.Text = uneStringDuree
            
            SpreadInfoVidage.Col = 6
            SpreadInfoVidage.Text = CLng(unParcours.maDistPar / 10)
            SpreadInfoVidage.Col = 7
            SpreadInfoVidage.Text = UBound(unParcours.monTabAbsRep)
            SpreadInfoVidage.Col = 8
            SpreadInfoVidage.Text = unParcours.monTypeMesure
        Next i
        'On vide les parcours issus du mtb
        ViderColParcours maColParcoursMTB
        Me.MousePointer = vbDefault
    Else
        'Cas d'erreur en lecture du MTB
        Tag = "A fermer"
    End If
End Sub
