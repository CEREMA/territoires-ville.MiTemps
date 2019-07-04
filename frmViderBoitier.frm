VERSION 5.00
Begin VB.Form frmViderBoitier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vider le boitier dans un fichier"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   Icon            =   "frmViderBoitier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame FrameInfoUser 
      Caption         =   "Messages d'information pour l'utilisateur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   9855
      Begin VB.TextBox TextInfo 
         Height          =   2895
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   360
         Width           =   9615
      End
   End
   Begin VB.Frame FrameParam 
      Caption         =   "Param�tres du transfert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   8655
      Begin VB.ComboBox ComboVit 
         Height          =   315
         ItemData        =   "frmViderBoitier.frx":0442
         Left            =   5160
         List            =   "frmViderBoitier.frx":0458
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox ComboCOM 
         Height          =   315
         ItemData        =   "frmViderBoitier.frx":047E
         Left            =   1920
         List            =   "frmViderBoitier.frx":0488
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bauds"
         Height          =   195
         Left            =   6120
         TabIndex        =   10
         Top             =   420
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vitesse de transmision : "
         Height          =   195
         Left            =   3360
         TabIndex        =   9
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Port de communication : "
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   1755
      End
   End
   Begin VB.Frame FrameFile 
      Caption         =   "Enregistrement dans le fichier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8655
      Begin VB.CommandButton btnParcourir 
         Caption         =   "Parcourir..."
         Default         =   -1  'True
         Height          =   375
         Left            =   7560
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TextNomFich 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   7335
      End
   End
   Begin VB.CommandButton btnAnnuler 
      Cancel          =   -1  'True
      Caption         =   "Interrompre"
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton btnTransferer 
      Caption         =   "Transf�rer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmViderBoitier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SectorsPerCluster As Long
Dim BytesPerSector As Long
Dim FreeClusters As Long
Dim TotalClusters As Long
Dim FreeBytes As Single
Dim TotalBytes As Single
Dim ArretTrans As Boolean
Dim TimeOut As Double

'****** D�finition des caract�res STX, ETX, ENQ, ACK, NAK *******
Private ENQ As String
Private STX As String
Private ETX As String
Private ACK As String
Private NAK As String
Private CR As String
Private LF As String
Private ETB As String

Public FichierOuvrir As String
Private TamponEntr�e As String

Public Function SommedeControle(A$)

    l = Len(A$)
    SommedeControle = Asc(Mid$(A$, 1, 1))
    For i% = 2 To l
        SommedeControle = SommedeControle + Asc(Mid$(A$, i%, 1))
    Next i%
    SommedeControle = SommedeControle And &H7F
End Function

Private Sub btnAnnuler_Click()
    ArretTrans = True
    TimeOut = 10
End Sub

Private Sub btnFermer_Click()
    'Fermeture de la feuille  *******************************
    If frmMain.MSComm1.PortOpen Then frmMain.MSComm1.PortOpen = False
    ArretTrans = True
    Close #1
    
    'Sauvegarde des param�tres de transmission pour la fois suivante
    SaveSetting App.Title, "Transmission", "Port", ComboCOM.Text
    SaveSetting App.Title, "Transmission", "Vitesse", ComboVit.Text
    
    Unload Me
End Sub

Private Sub btnParcourir_Click()
    'Ouvre la boite save avec read only masqu�
    FichierOuvrir = frmMain.ChoisirFichier(MsgSaveAs, MsgMtbFile, CurDir)
    If FichierOuvrir = "" Then Exit Sub
    
    TextNomFich.Text = frmMain.dlgCommonDialog.FileName
    FichierOuvrir = frmMain.dlgCommonDialog.FileName
    
'**************** V�rification de la place disponible sur l'unit� ******************

    GetDiskFreeSpace Left$(FichierOuvrir, 3), SectorsPerCluster, BytesPerSector, FreeClusters, TotalClusters
    FreeBytes = CSng(BytesPerSector) * SectorsPerCluster * FreeClusters
    TotalBytes = CSng(BytesPerSector) * SectorsPerCluster * TotalClusters
    
    If FreeBytes < 165000 Then
        MsgBox "Pas assez de place sur l'unit� de stockage : Corrigez et recommencer"
        Exit Sub
    End If
    
    TextInfo.Text = ""
    TextInfo.Text = "Sur le terminal, dans le menu d'accueil choisir '3 - R�sultats'" & CR & LF & "S�lectionner la m�me vitesse de transmission que ci-dessus et choisir sortie IBM-PC jusqu'� l'affichage du message 'ATTENTE COMMANDE'" & CR & LF & "Ensuite cliquer sur le bouton Transf�rer de la fen�tre pour continuer" & CR & LF
    btnTransferer.Enabled = True
End Sub

Private Sub btnTransferer_Click()
    Dim NumBloc As String * 1
    Dim CarEntr�e As String * 1
    Dim uneStringTmp As String
    
    On Error GoTo Erreur_port
    ArretTrans = False
    
    '*************** Ouverture du dialogue avec les param�tres d�finis *****************
    frmMain.MSComm1.PortOpen = True
    frmMain.MSComm1.SThreshold = 1
    frmMain.MSComm1.RThreshold = 1
    frmMain.MSComm1.InputLen = 1
    
    btnTransferer.Enabled = False
    Open FichierOuvrir For Output As #1


    '******************************* Message � envoyer *********************************
    A$ = ENQ & "0000,T" & ETX
    Sc = SommedeControle(A$)
    frmMain.MSComm1.Output = A$ & Chr$(Sc)
    
    '************************ Boucle d'attente de la fin d'envoi ***********************

    Do
    Loop Until frmMain.MSComm1.OutBufferCount = 0

    TextInfo.Text = TextInfo.Text & CR & LF & "Transfert en cours" & CR & LF
    'Pour �tre en fin de message ==> Scroll en bas
    TextInfo.SelStart = Len(TextInfo.Text)
    DoEvents
    

    Do
'Etiquette mis en commentaire, il servait en cas de NAK pour recommencer
'la lecture du bloc erron�, mais le NAK ne marche pas donc on ne fait rien
'd�but:
        TamponEntr�e = ""
        TimeOut = Timer + 10
        
        '*********************** Boucle de r�ception des caract�res ************************
        Do
            Do
                DoEvents
            Loop While frmMain.MSComm1.InBufferCount = 0 And Timer < TimeOut
            If ArretTrans Then GoTo Arret
            CarEntr�e = frmMain.MSComm1.Input
            TamponEntr�e = TamponEntr�e & CarEntr�e
        Loop Until CarEntr�e = ETB Or CarEntr�e = ETX Or Timer > TimeOut
    
        '*************** R�ception de la somme de contr�le du message **********************
        Do
            DoEvents
        Loop While frmMain.MSComm1.InBufferCount = 0 And Timer < TimeOut
        If ArretTrans Then GoTo Arret
        SdCR = frmMain.MSComm1.Input
            
        '**************** Calcul de la longueur du message (sortie par timeout ?) **********
        If Len(TamponEntr�e) < 3 Then
            MsgBox "Impossible d'�tablir le dialogue :" & CR & LF & "- v�rifier les connexions" & CR & LF & "- v�rifier que le boitier affiche ATENTE COMMANDE" & CR & LF & "Fermer la fen�tre et recommencer"
            GoTo Arret
        End If
         
        '********************* R�cup�ration du num�ro de bloc ******************************
        NumBloc = Mid$(TamponEntr�e, 5, 1)
            
        '**************** Calcul de la somme de contr�le de la r�ception *******************
        SdcC = SommedeControle(TamponEntr�e)
    
        '****************************** Envoi de ACK ou NAK ********************************
        If Chr$(SdcC) = SdCR Or (SdcC = 0 And SdCR = "") Then
            A$ = ACK & NumBloc
            frmMain.MSComm1.Output = A$
            '************ Boucle d'envoi **************
            Do
            Loop Until frmMain.MSComm1.OutBufferCount = 0
        Else
            A$ = NAK & NumBloc
            MsgBox "Erreur de transmission du bloc en cours" & CR & LF & "Fermer la fen�tre et recommencer"
            GoTo Arret
            'Mis en commentaire car le NAK n'est pas g�r�
            'frmMain.MSComm1.Output = A$
            '************ Boucle d'envoi **************
            'Do
            'Loop Until frmMain.MSComm1.OutBufferCount = 0
            'GoTo d�but
        End If
        FinMessage = InStr(6, TamponEntr�e, "," & LF)
        If FinMessage > 6 Then
            'Epuration des ent�tes d�but et fin du bloc transmis
            TamponEntr�e = Mid$(TamponEntr�e, 6, FinMessage - 5)
            'Compl�tude � 80 caract�res apr�s la derniere virgule
            If Len(TamponEntr�e) > 80 Then
                uneStringTmp = ""
            Else
                uneStringTmp = Space(80 - Len(TamponEntr�e))
            End If
            TamponEntr�e = TamponEntr�e + uneStringTmp
            'Ecriture dans le fichier mtb
            Print #1, TamponEntr�e;
            'Sans formatage � 80 caract�res
            'Print #1, Mid$(TamponEntr�e, 6, FinMessage - 5);
        Else
            MsgBox "Le message n'est pas correctement format�"
            GoTo Arret
        End If
    Loop Until CarEntr�e = ETX
    Close #1
    
    TextInfo.Text = TextInfo.Text & "Fin de Transfert" & CR & LF
    TextInfo.Text = TextInfo.Text & CR & LF & "Le contenu du boitier a �t� transfer� avec succ�s" & CR & LF
    'Pour �tre en fin de message ==> Scroll en bas
    TextInfo.SelStart = Len(TextInfo.Text)

    frmMain.MSComm1.PortOpen = False
    
    'Affichage d'un tableau de r�sultat des parcours lus du boitier
    Me.MousePointer = vbHourglass
    frmInfoVidage.Show vbModal, Me
    Me.MousePointer = vbDefault
    Exit Sub
    
Arret:
    frmMain.MSComm1.PortOpen = False
    TextInfo.Text = TextInfo.Text & "Arr�t de la transmission par l'op�rateur !" & CR & LF
    TextInfo.Text = TextInfo.Text & "Suppression du fichier " & FichierOuvrir
    'Pour �tre en fin de message ==> Scroll en bas
    TextInfo.SelStart = Len(TextInfo.Text)
    btnTransferer.Enabled = True
    Close #1
    'Text1.Text = ""
    
    '****************** Si n�cessaire effacement du fichier ****************************
    Kill FichierOuvrir
    '*************** ou mettre une boite deux options pour conserver *******************
    '**************** ou effacer le fichier au choix de l'op�rateur ********************
    Exit Sub

Erreur_port:
    If Err.Number = 8005 Then
        MsgBox "Erreur : " + Format(Err.Number) + Chr(13) + "Le port COM" + Format(frmMain.MSComm1.CommPort) + " est d�j� ouvert ou utilis� par un autre programme ou appareil.", vbCritical
    Else
        MsgBox "Erreur : " + Format(Err.Number) + Chr(13) + Err.Description, vbCritical
    End If
    btnTransferer.Enabled = True
    Exit Sub
End Sub

Private Sub ComboCOM_Click()
    frmMain.MSComm1.CommPort = ComboCOM.ListIndex + 1
End Sub


Private Sub ComboCOM_KeyPress(KeyAscii As Integer)
    ComboCOM_Click
End Sub

Private Sub ComboVit_Click()
    frmMain.MSComm1.Settings = ComboVit.Text & ",e,7,1"
End Sub

Private Sub ComboVit_KeyPress(KeyAscii As Integer)
    ComboVit_Click
End Sub

Private Sub Form_Load()
    'D�finition des caract�res STX, ETX, ENQ, ACK, NAK
    STX = Chr$(2)
    ETX = Chr$(3)
    ENQ = Chr$(5)
    CR = Chr$(13)
    LF = Chr$(10)
    ACK = Chr$(6)
    NAK = Chr$(&H15)
    ETB = Chr$(&H17)
    
    'R�cup�ration des param�tres de transmission
    '� partir de la base de registre
    'ou des valeurs par d�faut (COM1 et 9600 bauds)
    ComboCOM.Text = GetSetting(App.Title, "Transmission", "Port", "COM1")
    ComboVit.Text = GetSetting(App.Title, "Transmission", "Vitesse", "9600")
    
    'Affectation de ces param�tres de transfert
    frmMain.MSComm1.CommPort = ComboCOM.ListIndex + 1
    frmMain.MSComm1.Settings = ComboVit.Text & ",e,7,1"
    
    'Contexte d'aide
    HelpContextID = HelpID_WinViderBoitier
    
    CentrerFenetreEcran Me
End Sub

