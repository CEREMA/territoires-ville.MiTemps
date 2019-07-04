VERSION 5.00
Begin VB.Form frmRabouter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rabouter deux parcours de deux itin�raires diff�rents"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmRabouter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnAide 
      Caption         =   "Aide ou F1"
      Height          =   375
      Left            =   10160
      TabIndex        =   12
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton BtnAnnuler 
      Cancel          =   -1  'True
      Caption         =   "Fermer"
      Height          =   375
      Left            =   8765
      TabIndex        =   11
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton BtnRabouter 
      Caption         =   "Rabouter..."
      Height          =   375
      Left            =   7370
      TabIndex        =   10
      Top             =   6600
      Width           =   1215
   End
   Begin VB.ListBox ListParItiAval 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   11250
   End
   Begin VB.CommandButton BtnChoixItiAval 
      Caption         =   "Parcourir..."
      Height          =   375
      Left            =   10160
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox TextItiAval 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3600
      Width           =   9855
   End
   Begin VB.ListBox ListParItiAmont 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   11250
   End
   Begin VB.CommandButton BtnChoixItiAmont 
      Caption         =   "Parcourir..."
      Height          =   375
      Left            =   10160
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox TextItiAmont 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   9855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Liste des parcours du fichier itin�raire aval :"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   3045
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fichier itin�raire aval :"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Liste des parcours du fichier itin�raire amont :"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fichier itin�raire amont :"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "frmRabouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable stockant les rep�res des itin�raires amont et aval
Public maColRepAmont As New ColRepere
Public maColRepAval As New ColRepere
'Variable stockant les parcours des itin�raires amont et aval
Public maColParAmont As New ColParcours
Public maColParAval As New ColParcours
'Variable stockant les parcours et rep�res de l'itin�raire r�sultant de la fusion
Public maColParRes As New ColParcours
Public maColRepRes As New ColRepere

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

Private Sub btnAnnuler_Click()
    ViderColParcours maColParAmont
    ViderColParcours maColParAval
    ViderColRepere maColRepAmont
    ViderColRepere maColRepAval
    ViderColParcours maColParRes
    ViderColRepere maColRepRes
    Unload Me
End Sub

Private Sub BtnChoixItiAmont_Click()
    RemplirParcours maColRepAmont, maColParAmont, ListParItiAmont, TextItiAmont
End Sub

Private Sub BtnChoixItiAval_Click()
    RemplirParcours maColRepAval, maColParAval, ListParItiAval, TextItiAval
End Sub

Private Sub BtnRabouter_Click()
    Dim unNomFich As String, unNumLigNbPar As Integer
    Dim uneAbsMax As Long, uneAbsMin As Long, unFichId As Byte
    Dim unParAmont As Parcours, unParAval As Parcours
    Dim uneDistPar As Long, i As Long, unNbRepTop As Integer
    Dim unNbTopAmont As Integer, unNbTopAval As Integer
    Dim unTempsLastTopAmont As Long, unNumLastTopAmont As Long
    Dim unLastRepAmont As Repere, unFirstRepAval As Repere
    Dim unCoefEtaAmont As Single, unCoefEtaAval As Single
    Dim unTabAbsRep As Variant, unTabTmpRep As Variant
    
    If TextItiAmont.Text = "" Or TextItiAval.Text = "" Then
        MsgBox "Les fichiers itin�raires amont et aval doivent �tre renseigner.", vbCritical
    ElseIf ListParItiAmont.ListIndex = -1 Or ListParItiAval.ListIndex = -1 Then
        MsgBox "Il faut s�lectionner un parcours dans les fichiers itin�raires amont et aval.", vbCritical
    Else
        'R�cup�ration des parcours amont et aval
        Set unParAmont = maColParAmont(ListParItiAmont.ListIndex + 1)
        Set unParAval = maColParAval(ListParItiAval.ListIndex + 1)
        
        'Si les pas de mesure ne sont pas �gaux, on sort
        If unParAmont.monPasMesure <> unParAval.monPasMesure Then
            MsgBox "Le parcours amont et aval doivent avoir le m�me pas de mesure.", vbCritical
            Exit Sub
        End If
        
        'R�cup�ration du dernier rep�re du parcours amont,
        'la collection des rep�res n'est pas tri� par ordre croissant
        uneAbsMax = -100
        For i = 1 To maColRepAmont.Count
            If maColRepAmont(i).monAbsCurv > uneAbsMax Then
                uneAbsMax = maColRepAmont(i).monAbsCurv
                Set unLastRepAmont = maColRepAmont(i)
            End If
        Next i
        'R�cup�ration du premier rep�re du parcours aval,
        'la collection des rep�res n'est pas tri� par ordre croissant
        uneAbsMin = 10000000
        For i = 1 To maColRepAval.Count
            If maColRepAval(i).monAbsCurv < uneAbsMin Then
                uneAbsMin = maColRepAval(i).monAbsCurv
                Set unFirstRepAval = maColRepAval(i)
            End If
        Next i
        
        'Ajout du parcours fusionnant le deux parcours choisis dans le
        'fichier itin�raire r�sultant choisi
        unNomFich = frmMain.ChoisirFichier(MsgOpen, MsgMitFile, CurDir)
        If UCase(unNomFich) = UCase(TextItiAmont.Text) Or UCase(unNomFich) = UCase(TextItiAval.Text) Then
            MsgBox "Le fichier itin�raire o� le parcours rabout� est ajout� doit �tre diff�rent des fichiers itin�raires amont et aval.", vbCritical
        ElseIf unNomFich <> "" Then
            'Demande de confirmation � l'utilisateur en lui affichant le dernier
            'rep�re amont et le premier rep�re aval
            unMsgInfo = "Raboutage d'un parcours finissant au rep�re (Nom long = " + unLastRepAmont.monNomLong + " et Nom court = " + unLastRepAmont.monNomCourt + ")" + Chr(13)
            unMsgInfo = unMsgInfo + "avec un parcours commen�ant au rep�re (Nom long = " + unFirstRepAval.monNomLong + " et Nom court = " + unFirstRepAval.monNomCourt + ")"
            unMsgInfo = unMsgInfo + Chr(13) + Chr(13) + "Le parcours rabout� sera ajout� � la fin du fichier : " + unNomFich
            unMsgInfo = unMsgInfo + Chr(13) + Chr(13) + "Voulez-vous continuer ?"
            If MsgBox(unMsgInfo, vbYesNo + vbQuestion) = vbNo Then Exit Sub
            'D�but du traitement de raboutage
            Me.MousePointer = vbHourglass
            ViderColParcours maColParRes
            ViderColRepere maColRepRes
            'R�cup�ration du contenu de l'itin�raire (les parcours, les rep�res,...)
            'On retourne la ligne contenant les nombres de rep�res et de parcours
            unNumLigNbPar = RecupererContenuEtude(unNomFich, maColRepRes, maColParRes)
            If unNumLigNbPar > 0 Then
                'Pas d'erreur lors de la r�cup�ration/lecture,
                
                'Alimentation du parcours rabout� que l'on stockera dans la variable
                'globale monParToImport
                monParToImport.monNom = "Parcours rabout�"
                'R�cup�ration des donn�es communes
                monParToImport.monIsUtil = True
                monParToImport.maCouleur = unParAmont.maCouleur
                monParToImport.monEnqueteur = unParAmont.monEnqueteur
                monParToImport.monNumVeh = unParAmont.monNumVeh
                monParToImport.maMeteo = unParAmont.maMeteo
                monParToImport.maDate = unParAmont.maDate
                monParToImport.monJourSemaine = unParAmont.monJourSemaine
                monParToImport.monHeureDebut = unParAmont.monHeureDebut
                monParToImport.monPasMesure = unParAmont.monPasMesure
                monParToImport.monNumVeh = unParAmont.monNumVeh
                'Calcul du coefficient d'�talonnage
                'Si les deux coef d'�talonnage sont diff�rents on les met tous � 1
                'et on divise les tableaux d'interdistance par les coefs respectifs
                'en modifiant la valeur des coef d'�talonnage
                unCoefEtaAmont = unParAmont.monCoefEta
                unCoefEtaAval = unParAval.monCoefEta
                If Abs(unCoefEtaAval - unCoefEtaAmont) < 0.0000001 Then
                    monParToImport.monCoefEta = unParAmont.monCoefEta
                    unCoefEtaAmont = 1
                    unCoefEtaAval = 1
                Else
                    monParToImport.monCoefEta = 1
                    unCoefEtaAmont = unParAmont.monCoefEta
                    unCoefEtaAval = unParAval.monCoefEta
                End If
                'R�cup�ration du nombre de tops amont et aval
                unNbTopAmont = UBound(unParAmont.monTabAbsRep)
                unNbTopAval = UBound(unParAval.monTabAbsRep)
                'R�cup�ration du num�ro de pas de mesure du parcours amont
                'o� se trouve le dernier top amont
                unTempsLastTopAmont = unParAmont.monTabTempsRep(unNbTopAmont)
                unParAmont.DonnerInterDistance unTempsLastTopAmont, unNumLastTopAmont
                'Alimentation des pas de mesures, on enl�ve premier pas du parcours aval
                monParToImport.monNbPas = unNumLastTopAmont + unParAval.monNbPas - 1
                uneDistPar = 0
                For i = 1 To unNumLastTopAmont
                    monParToImport.monTabDist(i) = unParAmont.monTabDist(i) * unCoefEtaAmont
                    uneDistPar = uneDistPar + monParToImport.monTabDist(i)
                Next i
                For i = 2 To unParAval.monNbPas
                    monParToImport.monTabDist(i + unParAmont.monNbPas - 1) = unParAval.monTabDist(i) * unCoefEtaAval
                    uneDistPar = uneDistPar + monParToImport.monTabDist(i + unParAmont.monNbPas - 1)
                Next i
                monParToImport.monFirstPas = unParAmont.monFirstPas
                monParToImport.monLastPas = unParAval.monLastPas
                'Calcul de la distance parcourue = longueur du parcours
                monParToImport.maDistPar = uneDistPar
                'Calcul de la dur�e
                monParToImport.maDuree = unTempsLastTopAmont + unParAval.maDuree - unParAval.monFirstPas
                'Alimentation des rep�res top�s, somme des tops amont + aval
                'moins un, car celui faisant la jointure qui est top� deux fois
                unNbRepTop = unNbTopAmont + unNbTopAval - 1
                'Allocation dynamique des tableaux li�s aux rep�res top�s
                unTabAbsRep = monParToImport.monTabAbsRep
                unTabTmpRep = monParToImport.monTabTempsRep
                ReDim unTabAbsRep(1 To unNbRepTop)
                ReDim unTabTmpRep(1 To unNbRepTop)
                'Affectation pour chaque top
                For i = 1 To unNbTopAmont
                    unTabAbsRep(i) = unParAmont.monTabAbsRep(i) * unCoefEtaAmont
                    unTabTmpRep(i) = unParAmont.monTabTempsRep(i)
                Next i
                For i = 2 To unNbTopAval
                    unTabAbsRep(unNbTopAmont - 1 + i) = unTabAbsRep(unNbTopAmont) + unParAval.monTabAbsRep(i) * unCoefEtaAval
                    unTabTmpRep(unNbTopAmont - 1 + i) = unTempsLastTopAmont + unParAval.monTabTempsRep(i)
                Next i
                'Affectation des pointeurs sur le tableau
                'des abscisses curvilignes et des temps de passage des rep�res du parcours
                monParToImport.monTabAbsRep = unTabAbsRep
                monParToImport.monTabTempsRep = unTabTmpRep
                
                'Affichage des caract�ristiques du parcours fusionn� et
                'demande de confirmation d'ajout dans le fichier mit r�sultant
                unMsgInfo = "Caract�ristiques du parcours rabout� : "
                unMsgInfo = unMsgInfo + Chr(13) + "Nom = " + monParToImport.monNom
                unMsgInfo = unMsgInfo + Chr(13) + "Date = " + monParToImport.monJourSemaine + " " + Format(monParToImport.maDate) + " " + Format(monParToImport.monHeureDebut)
                unMsgInfo = unMsgInfo + Chr(13) + "Nb Tops = " + Format(UBound(monParToImport.monTabAbsRep))
                unMsgInfo = unMsgInfo + Chr(13) + "Distance au dernier Top = " + Format(CLng(monParToImport.monTabAbsRep(UBound(monParToImport.monTabAbsRep)) * monParToImport.monCoefEta / 10)) + " m"
                unMsgInfo = unMsgInfo + Chr(13) + "Distance parcourue = " + Format(CLng(monParToImport.maDistPar * monParToImport.monCoefEta / 10)) + " m"
                unMsgInfo = unMsgInfo + Chr(13) + "Dur�e = " + FormatterTempsEnHMNS(monParToImport.maDuree)
                unMsgInfo = unMsgInfo + Chr(13) + "Coef d'�talonnage = " + Format(monParToImport.monCoefEta)
                unMsgInfo = unMsgInfo + Chr(13) + Chr(13) + "Le parcours rabout� sera ajout� � la fin du fichier : " + unNomFich
                unMsgInfo = unMsgInfo + Chr(13) + Chr(13) + "Voulez-vous continuer ?"
                If MsgBox(unMsgInfo, vbYesNo + vbQuestion) = vbYes Then
                    ' Active la routine de gestion d'erreur.
                    'MsgBox "Suppression du On Error GoTo ErreurAjoutIti"
                    On Error GoTo ErreurAjoutIti
                    'Ouverture du fichier itin�raire charg� en mode ajout pour ajouter
                    'le parcours fusionn� que l'on stockera dans la variable globale
                    'monParToImport
                    unFichId = FreeFile(0)
                    Open unNomFich For Append As #unFichId
                    'Ecriture du parcours rabout� en fin du fichier mit r�sultant
                    EcrireDonneesParcoursDansFichierMIT unFichId, monParToImport
                    'On se met sur la ligne o� se trouve le nombre de rep�res et de parcours
                    'pour ajouter +1 au nombre de parcours en le formattant sur 4 caract�res
                    'pour �crire juste dans sa place (4 caract�res lors de la sauvegarde du
                    'fichier MIT).
                    'Cette position est donn�e par unNumLigNbPar d�finie plus haut dans
                    'cette fonction
                    Seek #unFichId, unNumLigNbPar
                    unTextLine = Format(maColRepRes.Count) + "," + Format(maColParRes.Count + 1, "000#")
                    Print #unFichId, unTextLine
                    'Fermeture du fichier mit r�sultat
                    Close #unFichId
                    MsgBox "Raboutage des parcours r�ussi.", vbInformation
                End If
            End If
            Me.MousePointer = vbDefault
        End If
    End If
    
    'Sortie pour �viter la gestion d'erreur
    On Error GoTo 0
    Exit Sub
    
    ' Routine de gestion d'erreur qui �value le num�ro d'erreur.
ErreurAjoutIti:
    
    ' Traite les autres situations ici...
    unMsg = MsgOpenError + unNomFich + Chr(13) + Chr(13) + MsgErreur + Format(Err.Number) + " : " + Err.Description
    If Err.Number = 70 Then
        unMsg = unMsg + " (" + UCase(MsgDejaOpen) + ")"
    End If
    MsgBox unMsg, vbCritical
    'Fermeture et r�ouverture en mode verrouill�
    Close #unFichId
    Me.MousePointer = vbDefault
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    Exit Sub
End Sub

Private Sub Form_Load()
    CentrerFenetreEcran Me
    'Contexte d'aide
    HelpContextID = HelpID_WinRabouter
End Sub

Public Sub RemplirParcours(uneColRep As ColRepere, uneColPar As ColParcours, uneListBox As ListBox, unTextFichIti As TextBox)
    Dim unNomFich As String, unNumLigNbPar As Integer

    unNomFich = frmMain.ChoisirFichier(MsgOpen, MsgMitFile, CurDir)
    If unNomFich <> "" Then
        Me.MousePointer = vbHourglass
        ViderColParcours uneColPar
        ViderColRepere uneColRep
        'R�cup�ration du contenu de l'�tude (les parcours, les rep�res,...)
        'On retourne la ligne contenant les nombres de rep�res et de parcours
        unNumLigNbPar = RecupererContenuEtude(unNomFich, uneColRep, uneColPar)
        If unNumLigNbPar Then
            'Pas d'erreur lors de la r�cup�ration/lecture,
            'donc on afiche la liste des parcours dans la liste d�roulante
            'ListParItiAmont
            RemplirListeBox uneListBox, uneColPar
            'Affichage du nom du fichier itin�raire
            unTextFichIti.Text = unNomFich
            'Affichage de la fin du nom de fichier si d�passe zone texte
            'en pla�ant le curseur en fin de texte
            unTextFichIti.SelStart = Len(unTextFichIti.Text)
        Else
            'Affichage du nom du fichier itin�raire
            unTextFichIti.Text = ""
            ViderColParcours uneColPar
            ViderColRepere uneColRep
            uneListBox.Clear
        End If
        Me.MousePointer = vbDefault
    End If
End Sub

Public Sub RemplirListeBox(uneListBox As ListBox, uneColPar As ColParcours)
    Dim unPar As Parcours, unNbTop As Integer, unNomPar As String * 20
    
    uneListBox.Clear
    For i = 1 To uneColPar.Count
        Set unPar = uneColPar(i)
        unNbTop = UBound(unPar.monTabAbsRep)
        'uneListBox.AddItem unPar.monNom + " | " + unPar.monJourSemaine + " " + Format(unPar.maDate) + " " + Format(unPar.monHeureDebut) + " Cf �talonnage = " + Format(unPar.monCoefEta, "0.0000") + " Dernier Top = " + Format(unPar.monTabAbsRep(unNbTop)) + " m " + Format(unNbTop) + " rep�res top�s"
        unEsp = Space(1)
        If Len(unPar.monNom) >= 20 Then
            unNomPar = Mid(LCase(unPar.monNom), 1, 20)
        Else
            unNomPar = LCase(unPar.monNom) + String(20 - Len(unPar.monNom), " ")
        End If
        uneListBox.AddItem unNomPar + unEsp + LCase(unPar.monJourSemaine) + String(8 - Len(unPar.monJourSemaine), " ") + " " + Format(unPar.maDate) + " " + Format(unPar.monHeureDebut) + unEsp + " pas = " + FormatterEnNCarLeft(4, Format(unPar.monPasMesure) + "s") + unEsp + "Nb.Top " + FormatterEnNCarLeft(3, CLng(unNbTop)) + unEsp + "dernier � " + Format(CLng(unPar.monTabAbsRep(unNbTop) / 10)) + " m"
    Next i
End Sub
