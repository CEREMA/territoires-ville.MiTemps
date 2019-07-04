Attribute VB_Name = "ModuleFichier"
Public Function LireFichierMTB(unNomFich As String) As Boolean
    'Lecture du fichier MTB pass� en param�tre
    Dim uneString As String, unMsg As String
    Dim uneStrTmp As String, unFichId As Byte
    Dim unePos As Integer, unNbRepTop As Integer, unNbPas As Long
    Dim unNbData As Variant, unParcours As Parcours
    Dim unTabTempsRep As Variant, unTabAbsRep As Variant
    Dim unNbParLu As Integer
    Dim unTabIdRep As Variant, unFormatMTB As Byte
    Dim unePartEnt As Integer, unePartDec As Integer
    
    'Test du format du fichier MTB
    unFormatMTB = IsNewFichierMTB(unNomFich)
    If unFormatMTB = BadMTB Then
        'Cas d'un mauvais format, on sort sans rien faire
        LireFichierMTB = False
        Exit Function
    End If
    
    ' Active la routine de gestion d'erreur.
    LireFichierMTB = True
    On Error GoTo ErreurLireMTB
    
    ' Ouvre le fichier en lecture.
    unFichId = FreeFile(0)
    Open unNomFich For Input As #unFichId
    
    'Lecture du fichier MTB
    'Effectue la boucle jusqu'� la fin du fichier.
    unNbData = 0
    unNbParLu = 1
    Do While Not EOF(unFichId)
        ' Lit les donn�es dans une variable.
        Input #unFichId, uneString
        'Incr�mentation du nombre de donn�es pour ce parcours
        unNbData = unNbData + 1
        unePos = InStr(1, uneString, "!", vbTextCompare)
        If unePos > 0 Then
            'Cas du premier ! trouv� (code pour lire les vieux fichiers MTB contenant des !!)
            Input #unFichId, uneStrTmp 'passage au ! suivant
            Input #unFichId, uneStrTmp 'Lecture du reste de la donn�e coup� par les !!
            'Reconstruction de la valeur globale
            If uneString = "!" Then
                uneString = uneStrTmp
            Else
                uneString = Mid(uneString, 1, Len(uneString) - 1) + uneStrTmp
            End If
        End If

        'Remplissage des donn�es du parcours
        Select Case unNbData  ' �value unNbData.
            Case 1
                If uneString <> "" Then
                    'Cr�ation du parcours issu du MTB
                    'avec affectation du nom et d'une couleur par d�faut
                    'Affectation d'une couleur par d�faut, on commence � 9
                    'pour �viter le gris (cf aide sur fonction QBColor)
                    Set unParcours = maColParcoursMTB.Add(uneString, QBColor(9 + unNbParLu Mod 6))
                    'Mise � z�ro du pas stockant le d�but d'un arr�t
                    unNumPasDebArret = 0
                End If
                'Si le fichier MTB a un 0D0A en fin de fichier, la derni�re lecture
                'donne une chiane vide pour le nom, c'est ainsi que l'on rep�re
                'ces fichiers MTB qui ont �t� sauv� par un �diteur DOS ou Windows
            Case 2
                unParcours.monEnqueteur = uneString
            Case 3
                unParcours.monNumVeh = uneString
            Case 4
                unParcours.maMeteo = CByte(uneString)
            Case 5
                unParcours.monJourSemaine = uneString
            Case 6
                'Remplacement des deux : par deux / des dates du MTB
                unePos = InStr(1, uneString, ":")
                uneString = Mid(uneString, 1, unePos - 1) + "/" + Mid(uneString, unePos + 1)
                unePos = InStr(1, uneString, ":")
                uneString = Mid(uneString, 1, unePos - 1) + "/" + Mid(uneString, unePos + 1)
                unParcours.maDate = DateValue(uneString)
            Case 7
                'On enl�ve les centi�mes de secondes de l'heure de d�part
                'Date VB va jusqu'au seconde et le spread aussi
                'On supprime les trois derniers caract�res hh:MM:ss:mm
                'donc :mm
                unParcours.monHeureDebut = TimeValue(Mid(uneString, 1, Len(uneString) - 3))
            Case 8
                'Lecture du type de mesure (1 car = D presque toujours)
                unParcours.monTypeMesure = uneString
            Case 9
                '9 (= lecture du nbre de rep�res th�oriques) on ne fait rien
                '==> on passe � la lecture suivante
                uneString = uneString
            Case 10
                'On remplace le caract�re d�cimale en cours par le point
                'sinon CSng plante, en effet le s�parateur est le point
                'dans les fichiers MTB
                unePosSepDec = InStr(1, uneString, ".")
                'R�cup de la partie enti�re � gauche du point d�cimale
                unePartEnt = Format(Mid(uneString, 1, unePosSepDec - 1))
                'R�cup de la partie d�cimale (4 chiffres maxi) � gauche du point d�cimale
                unePartDec = Format(Mid(uneString, unePosSepDec + 1))
                unParcours.monCoefEta = unePartEnt + 0.0001 * unePartDec
            Case 11
                unParcours.monPasMesure = CLng(uneString)
            Case 12
                unParcours.maDuree = CLng(uneString)
                unParcours.monTDebSection = 0
                unParcours.monTFinSection = CLng(uneString)
            Case 13
                unParcours.maDistPar = CLng(uneString) / unParcours.monCoefEta
                unParcours.maDistParSection = CLng(uneString) / unParcours.monCoefEta
                'On divise car les donn�es sont d�j� multipli�s par le coef d'�talonnage
                'Ainsi les distances stock�es sont ind�pendants du coef d'�talonnage
            Case 14
                'Stockage du nombre de pas de mesure
                unNbPas = CLng(uneString)
                unParcours.monNbPas = unNbPas
                If unNbPas > NbPasMax Then
                    'Cas ou le nombre de mesures d�passe le nb maxi
                    'fix�, on sort sans rien faire
                    MsgBox "Le nombre de mesures, valant " + Format(unNbPas) + ", du parcours " + Format(unParcours.monNom) + " d�passe le nombre de mesures maximun fix� � " + Format(NbPasMax), vbCritical
                    LireFichierMTB = False
                    Exit Function
                End If
            Case 15
                unParcours.monFirstPas = CInt(uneString)
            Case 16
                unParcours.monLastPas = CInt(uneString)
            Case 17
                'Stockage du nombre de rep�res top�s
                unNbRepTop = CInt(uneString)
                If unNbRepTop < 2 Then
                    'Obligation d'avoir au moins deux rep�res
                    MsgBox "Le nombre de rep�res top�s du parcours lu en position " + Format(unNbParLu) + " devrait �tre sup�rieur ou �gal � 2", vbInformation
                    'on se sort plus car cela ne g�ne en rien le reste et en plus
                    'on peut continuer de lire les parcours suivants
                    'LireFichierMTB = False
                    'Exit Function
                End If
                'Allocation dynamique des tableaux li�s aux rep�res top�s
                unTabAbsRep = unParcours.monTabAbsRep
                ReDim unTabAbsRep(1 To unNbRepTop)
                unTabTempsRep = unParcours.monTabTempsRep
                ReDim unTabTempsRep(1 To unNbRepTop)
                If unFormatMTB = NewMTB Then
                    'Cas du nouveau format de fichier
                    unTabIdRep = unParcours.monTabIdRep
                    ReDim unTabIdRep(1 To unNbRepTop)
                End If
            Case Else
                'Autres valeurs
                If unNbData = 17 + unNbPas + unNbRepTop * 2 Then
                    'Test de Fin de lecture des donn�es de ce parcours
                    'Lecture du temps de passage du dernier rep�re
                    unTabTempsRep(unNbRepTop) = CLng(uneString)
                    
                    If unFormatMTB = NewMTB Then
                        'Cas du nouveau format de fichier
                        'On lit les champs d'identification des rep�res
                        For i = 1 To unNbRepTop
                            Input #unFichId, uneString
                            unTabIdRep(i) = uneString
                        Next i
                        'Affectation des pointeurs sur le tableau
                        'des champs d'identification des rep�res du parcours
                        unParcours.monTabIdRep = unTabIdRep
                    End If
                    
                    'On remet unNbData � 0 pour la lecture des donn�es
                    'des parcours suivants �ventuelles
                    unNbData = 0
                
                    'Incr�mentation du nombre de parcours
                    unNbParLu = unNbParLu + 1
                    
                    'Affectation des pointeurs sur les tableaux du parcours
                    unParcours.monTabTempsRep = unTabTempsRep
                    unParcours.monTabAbsRep = unTabAbsRep
                ElseIf unNbData > 17 + unNbPas + unNbRepTop Then
                    'Lecture des temps de passages des rep�res.
                    unTabTempsRep(unNbData - 17 - unNbPas - unNbRepTop) = CLng(uneString)
                ElseIf unNbData > 17 + unNbPas Then
                    'Lecture des abscisses des rep�res.
                    unTabAbsRep(unNbData - 17 - unNbPas) = CLng(uneString) / unParcours.monCoefEta
                    'On divise car les donn�es sont d�j� multipli�s par le coef d'�talonnage
                    'Ainsi les abscisses stock�es sont ind�pendants du coef d'�talonnage
                ElseIf unNbData > 17 Then
                    'Lecture des distances parcourues par pas
                    unParcours.monTabDist(unNbData - 17) = CLng(uneString) / unParcours.monCoefEta
                    'On divise car les donn�es sont d�j� multipli�s par le coef d'�talonnage
                    'Ainsi les abscisses stock�es sont ind�pendants du coef d'�talonnage
                    
                    'Calcul du nombre et du temps d'arr�ts
                    If CLng(uneString) = 0 Then
                        'Cas o� on est arr�t�
                        If unNumPasDebArret = 0 Then
                            'Stockage du pas d�butant l'arr�t, tjs > 0
                            unNumPasDebArret = unNbData - 17
                            unParcours.monNbArret = unParcours.monNbArret + 1
                        End If
                        'Cumul du temps d'arr�t en dixi�me de seconde,
                        'le pas de mesure est en seconde
                        If unNbData - 17 = 1 Then
                            unParcours.monTpsArret = unParcours.monTpsArret + unParcours.monFirstPas
                        ElseIf unNbData - 17 = unParcours.monNbPas Then
                            unParcours.monTpsArret = unParcours.monTpsArret + unParcours.monLastPas
                        Else
                            unParcours.monTpsArret = unParcours.monTpsArret + unParcours.monPasMesure * 10
                        End If
                    Else
                        'Cas o� on n'est pas arr�t�
                        'Remise � z�ro du pas stockant le d�but de l'arr�t
                        unNumPasDebArret = 0
                    End If
                End If
        End Select
    Loop
    
    'Ferme le fichier
    Close #unFichId
    
    'D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    'Sortie de la proc�dure pour �viter le passage
    'dans la gestion d'erreur
    Exit Function

ErreurLireMTB:
    'Cas o� erreur de lecture du fichier MTB
    If unNbData = 0 And unNbParLu > maColParcoursMTB.Count Then
        'Lecture en fin de fichier normal, ce n'est pas une erreur
        'les mtb finissent tous par une virgule, on a lu le dernier champ
        'le nbdata remis � z�ro mais le nbParLu incr�mente avant le add dans
        'la collection des parcours, donc c'est le cas o� on lit apr�s la fin
        'du fichier, ap�s la derni�re virgule
        LireFichierMTB = True
    Else
        'Cas des erreurs
        MsgBox MsgErreur + Format(Err.Number) + " " + MsgIn + "LireFichierMTB / " + Err.Description, vbCritical
        ViderColParcours maColParcoursMTB
        LireFichierMTB = False
    End If
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    'Ferme le fichier
    Close #unFichId
    Exit Function
End Function

Public Function IsNewFichierMTB(unNomFich As String) As Byte
    'Function indiquant si le fichier MTB pass� en param�tre
    'est au nouveau format de fichier ou pas
    'Le nouveau format comporte des champs d'identification de rep�re
    '(1 caract�re par rep�re) en plus
    'Valeur de retour :
    '   OldMTB si vieux format
    '   NewMTB si nouveau format
    '   BadMTB si mauvais format de fichier ( tous les autres cas)
    'Ces constantes sont d�finies dans le module ModuleMain
    
    Dim uneString As String, unMsg As String
    Dim uneStrTmp As String, unFichId As Byte
    Dim unePos As Integer, unNbRepTop As Integer, unNbPas As Long
    Dim unNbData As Variant, uneFinLecture As Boolean
    
    ' Active la routine de gestion d'erreur.
    On Error GoTo ErreurNewMTB
    
    ' Ouvre le fichier en lecture.
    unFichId = FreeFile(0)
    Open unNomFich For Input As #unFichId
    
    'Initialisation
    IsNewFichierMTB = NewMTB
    
    'Lecture du fichier MTB
    'Effectue la boucle jusqu'� la fin du fichier.
    unNbData = 0
    uneFinLecture = False
    Do While Not uneFinLecture
        ' Lit les donn�es dans une variable.
        Input #unFichId, uneString
        unNbData = unNbData + 1
        unePos = InStr(1, uneString, "!", vbTextCompare)
        If unePos > 0 Then
            'Cas du premier ! trouv� (code pour lire les vieux fichiers MTB contenant des !!)
            Input #unFichId, uneStrTmp 'passage au ! suivant
            Input #unFichId, uneStrTmp 'Lecture du reste de la donn�e coup� par les !!
            'Reconstruction de la valeur globale
            If uneString = "!" Then
                uneString = uneStrTmp
            Else
                uneString = Mid(uneString, 1, Len(uneString) - 1) + uneStrTmp
            End If
        End If

        'R�cup�ration de certaines donn�es du parcours
        If unNbData = 14 Then
            'Stockage du nombre de pas de mesure
            unNbPas = CLng(uneString)
        ElseIf unNbData = 17 Then
            'Stockage du nombre de rep�res top�s
            unNbRepTop = CInt(uneString)
        ElseIf unNbData = 17 + unNbPas + unNbRepTop * 2 Then
            'Test de Fin de lecture des donn�es de ce parcours
            'Lecture des unNbRepTop champs suivants pour voir si c'est
            'un nouveau de fichier avec les identifications de rep�res
            'sur un caract�re.
            'Si ces unNbRepTop champs ont tous une longueur <= 1
            'alors nouveau format de fichier sinon ancien fichier
            'Si ancien fichier et pas d'autres parcours,
            '==> le champ lu suivant d�clenche EOF
            For i = 1 To unNbRepTop
                Input #unFichId, uneString
                'Si on d�passe la fin du fichier ==> Vieux format fichier
                'on part dans la gestion d'erreur ErreurNewMTB plus bas
                'Si la longueur chaine est > 1 ==> Vieux format de fichier,
                'pas de champ identification des rep�res
                If Len(uneString) > 1 Then
                    IsNewFichierMTB = OldMTB
                    Exit For
                End If
            Next i
            
            'Arr�t de la boucle de lecture
            uneFinLecture = True
        End If
    Loop
    
    'Ferme le fichier
    Close #unFichId
    
    'D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    'Sortie de la proc�dure pour �viter le passage
    'dans la gestion d'erreur
    Exit Function

ErreurNewMTB:
    'Cas o� erreur de lecture du fichier MTB
    If unNbData = 17 + unNbPas + unNbRepTop * 2 Then
        'Lecture en fin de fichier normal, ce n'est pas une erreur
        'c'est un fichier MTB au vieux format
        IsNewFichierMTB = OldMTB
    Else
        'Cas des autres erreurs
        MsgBox MsgErreur + Format(Err.Number) + " " + MsgIn + "IsNewFichierMTB / " + Err.Description, vbCritical
        IsNewFichierMTB = BadMTB
    End If
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    'Ferme le fichier
    Close #unFichId
    Exit Function
End Function

Public Sub OuvrirEtude(unNomFich As String)
    'Ouvre l'�tude contenue dans le fichier pass� en param�tre
    Dim unRep As Repere, unBool As Boolean, unBool2 As Boolean
    Dim unMinD As Single, unMaxD As Single
    Dim unPar As Parcours, unParMoyen As Parcours
    Dim uneAbsCurv As Long, unTypeIco As Byte, unY1 As Long, unY2 As Long
    Dim unNbPas As Long, unNbRepTop As Long, j As Long
    Dim unNbRep As Integer, unNbPar As Integer
    Dim unLong1 As Long, unLong2 As Long, unLong3 As Long, unByte As Byte
    Dim unInt1 As Integer, unInt2 As Integer, unInt3 As Integer
    Dim uneString As String, uneString1 As String, uneString2 As String
    Dim uneString3 As String, uneHeure As Date, uneDate As Date
    Dim frmD As frmDocument, unePos As Integer
    Dim unTabAbsRep As Variant, unTabTempsRep As Variant
    Dim unTabSng(1 To NbPasMax) As Single 'Pour stocker les lectures de single
    Dim unCheckSection As Integer, unIndRepDeb As Integer, unIndRepFin As Integer
    
    'Si protection invalide on ne fait rien ancien syst�me neutralis�
    'If ProtectCheck(2) <> 0 Then Exit Sub
    
    'Lecture du fichier .mit
    ' Active la routine de gestion d'erreur.
    'MsgBox "Suppression du On Error GoTo ErreurLecture"
    On Error GoTo ErreurLecture
    
    'Ouverture du fichier en lecture lock�e pour �viter deux ouvertures
    unFichId = FreeFile(0)
    Open unNomFich For Input Lock Read Write As #unFichId
    
    'Cr�ation de la nouvelle fen�tre qui contiendra l'�tude
    Set frmD = New frmDocument
    frmD.Visible = False 'pour corriger un bug bizarre VB
    'Titre de la fen�tre itin�raire
    frmD.Caption = MsgIti0 + unNomFich
    
    'Stockage du fichier id
    frmD.monFichId = unFichId
    
    'Lecture de l'ent�te des fichiers *.mit
    Input #unFichId, uneString
    If uneString <> "Fichier " + App.Title Then
        'Cas d'un fichier qui n'est pas un fichier MiTemps
        '===> Fermeture du fichier.
        Close #unFichId
        MsgBox MsgErreur + MsgFileNotFile + App.Title + " version " + Format(App.Major) + "." + Format(App.Minor), vbCritical
        ' D�sactive la r�cup�ration d'erreur.
        On Error GoTo 0
        Exit Sub
    Else
        'Cas d'un fichier MiTemps *.Mit de la version 3.0
        '1�re ligne du fichier MIT = "Fichier MiTemps"
        
        'R�cup�ration des libell�s des conditions m�t�o
        Input #unFichId, uneString
        For i = 0 To 7
            unePos = InStr(1, uneString, ",")
            frmD.maColMeteo.Add Mid(uneString, 1, unePos - 1)
            uneString = Mid(uneString, unePos + 1)
        Next i
        'Dernier libell�
        frmD.maColMeteo.Add uneString
        
        'R�cup�ration des min et max total en distance (parcours complet sans section)
        Input #unFichId, unTabSng(1), unTabSng(2)
        frmD.monMinDtot = unTabSng(1)
        frmD.monMaxDtot = unTabSng(2)
        'R�cup�ration des min et max en distance, vitesse et temps
        'Input #unFichId, unTabSng(1), unTabSng(2), unTabSng(3), unTabSng(4), unTabSng(5), unTabSng(6)
        Input #unFichId, unMinD, unMaxD, unTabSng(3), unTabSng(4), unTabSng(5), unTabSng(6)
        frmD.monMinD = unMinD 'unTabSng(1)
        frmD.monMaxD = unMaxD 'unTabSng(2)
        frmD.monMinV = unTabSng(3)
        frmD.monMaxV = unTabSng(4)
        frmD.monMinT = unTabSng(5)
        frmD.monMaxT = unTabSng(6)
        
        'R�cup�ration du nom de l'itin�raire et de sa longueur
        Input #unFichId, uneString, uneString1
        frmD.TextNomIti.Text = uneString
        frmD.TextLongIti.Text = uneString1
        frmD.maLongIti = CLng(uneString1)
        
        'R�cup�ration des donn�es de la section de travail
        Input #unFichId, unCheckSection, unIndRepDeb, unIndRepFin
        
        'On remet � z�ro le tableau de rep�res
        frmD.SpreadRepere.MaxRows = 0
        'R�cup�ration du nombre de rep�res et de parcours
        Input #unFichId, unNbRep, unNbPar
        'R�cup�ration et cr�ation des rep�res
        For i = 1 To unNbRep
            Input #unFichId, uneString, uneString1, uneAbsCurv, unTypeIco
            Set unRep = CreerRepere(frmD, uneString, uneString1, uneAbsCurv, unTypeIco)
        Next i
        'Remise � z�ro de unNbRep
        unNbRep = 0
        
        'R�cup�ration des donn�es des parcours, pour chaque parcours :
        'Ligne 1 : Les champs modifiables, Ligne 2 : les non-modifiables
        'stock�s dans le spread parcours dont le nombre de rep�res top�s
        'Ligne 3 : le nombre de pas de mesure,
        '          le pas de mesure en secondes,
        '          le premier et dernier pas de mesure
        'Ligne 4 � nb pas  modulo 13, (13 pas de mesure par ligne)
        'Apr�s les N absCurv des rep�res top�s (10 abscurv par ligne)
        'Puis les N Temps passage des rep�res top�s (10 temps par ligne)
        For i = 1 To unNbPar
            Input #unFichId, uneString, unBool, unBool2, unLong1, uneString1, uneString2, unByte, uneDate, uneString3, uneHeure
            Set unPar = frmD.maColParcours.Add(uneString, unLong1)
            unPar.monIsUtil = unBool
            unPar.monIsParcoursMoyen = unBool2
            unPar.monEnqueteur = uneString1
            unPar.monNumVeh = uneString2
            unPar.maMeteo = unByte
            unPar.maDate = uneDate
            unPar.monJourSemaine = uneString3
            unPar.monHeureDebut = uneHeure
            
            Input #unFichId, unNbRepTop, unLong1, unLong2, unLong3, unTabSng(1)
            unPar.maDistPar = unLong1
            unPar.maDuree = unLong3
            unPar.monCoefEta = unTabSng(1)
            
            'Stockage dans unNbRep du nombre de rep�res maxis
            If unNbRepTop > unNbRep Then unNbRep = unNbRepTop
            
            Input #unFichId, unNbPas, unLong1, unInt1, unInt2
            unPar.monNbPas = unNbPas
            unPar.monPasMesure = unLong1
            unPar.monFirstPas = unInt1
            unPar.monLastPas = unInt2
                        
            'R�cup�ration des pas de mesure
            For j = 1 To unNbPas
                'Lecture de la donn�e
                Input #unFichId, unTabSng(1)
                unPar.monTabDist(j) = unTabSng(1)
            Next j
                        
            'Allocation dynamique des tableaux li�s aux rep�res top�s
            unTabAbsRep = unPar.monTabAbsRep
            unTabTempsRep = unPar.monTabTempsRep
            ReDim unTabAbsRep(1 To unNbRepTop)
            ReDim unTabTempsRep(1 To unNbRepTop)
            
            'R�cup�ration des abs curv des rep�res top�s
            For j = 1 To unNbRepTop
                'Lecture de la donn�e
                Input #unFichId, unLong1
                unTabAbsRep(j) = unLong1
            Next j
        
            'R�cup�ration des temps de passage des rep�res top�s
            For j = 1 To unNbRepTop
                'Lecture de la donn�e
                Input #unFichId, unTabSng(1)
                unTabTempsRep(j) = unTabSng(1)
            Next j
            
            'Affectation des pointeurs sur les tableaux du parcours
            unPar.monTabAbsRep = unTabAbsRep
            unPar.monTabTempsRep = unTabTempsRep
        Next i
    
        If unNbPar > 0 Then
            If frmD.maColParcours(1).monIsParcoursMoyen = False Then
                'Cas o� le premier parcours n'est pas le parcours moyen
                '==> il faut le cr�er et le calculer
                'Cela se produit uniquement si on importe des parcours dans
                'un fichier itin�raire n'ayant pas de parcours ou n'ayant jamais
                '�t� ouvert puis sauvegarder (car le parcours moyen est alors cr��)
                
                'Cr�ation du parcours moyen et
                'Ajout en t�te dans les parcours du nouvel itin�raire
                'Le parcours moyen sera toujours celui en premi�re position
                ' qbcolor(0) = noir
                Set unPar = New Parcours
                frmD.maColParcours.Ajouter unPar, True
                'Indication que ce parcours cr�� est le parcours moyen
                'et qu'il sera utilis�
                unPar.monIsParcoursMoyen = True
                unPar.monIsUtil = True
            Else
                Set unPar = frmD.maColParcours(1)
            End If
            
            'Allocation dynamique des tableaux li�s aux rep�res top�s
            'Le parcours moyen aura un nombre top valant le plus grand nombre
            'de top de tous les autres parcours
            unTabAbsRep = unPar.monTabAbsRep
            unTabTempsRep = unPar.monTabTempsRep
            ReDim unTabAbsRep(1 To unNbRep)
            ReDim unTabTempsRep(1 To unNbRep)
            'Affectation des pointeurs sur les tableaux du parcours
            unPar.monTabAbsRep = unTabAbsRep
            unPar.monTabTempsRep = unTabTempsRep
            
            'Calcul des caract�ristiques g�n�rales
            'et du tableau de distances parcourues par pas du parcours
            If unCheckSection = 0 Then
                'Stockage des abs d�but et fin du parcours
                unY1 = -100
                unY2 = 1000000
            Else
                'Stockage des abs d�but et fin de la section de travail du parcours
                unY1 = frmD.maColRepere(unIndRepDeb).monAbsCurv
                unY2 = frmD.maColRepere(unIndRepFin).monAbsCurv
            End If
            ActualiserParcoursMoyen frmD.maColParcours(1), frmD.maColParcours, unY1, unY2
        End If
        
        'Remplissage du spread parcours de la fen�tre itin�raire
        RemplirSpreadParcours frmD
        'Remplissage du spread parcours avec les libell�s m�t�o
        'de la fen�tre itin�raire
        RemplirMeteoSpreadParcours frmD
    End If
        
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
        
    'Mettre � jour liste des fichiers r�cents
    ActualiserListeFichiersRecents unNomFich
    
    'Vidage de tous les events en attente
    DoEvents
    'Mise en zoom total pour �viter les redessins successifs lors de la
    'cr�ation des rep�res
    frmD.monMinD = frmD.monMinDtot
    frmD.monMaxD = frmD.monMaxDtot
    'Affichage de la fen�tre qui d�clenche le dessin des rep�res
    frmD.Show
    'Remise de l'englobant stock� lors de la sauvegarde du fichier mit
    frmD.monMinD = unMinD
    frmD.monMaxD = unMaxD
        
    'Mise � jour des donn�es de la section de travail car c'est CreerRepere
    'qui cr�e les �l�ments des combobox d�but et fin de section
    frmD.ComboRepDebSec.ListIndex = unIndRepDeb - 1
    frmD.ComboRepFinSec.ListIndex = unIndRepFin - 1
    'Stockage dans le tag des derni�res valeurs valides
    frmD.ComboRepDebSec.Tag = frmD.ComboRepDebSec.Text
    frmD.ComboRepFinSec.Tag = frmD.ComboRepFinSec.Text
    frmD.CheckSection.Value = unCheckSection
    
    'Appeler du click event de CheckSection pour recalcul l'englobant et
    'les vitesses, temps d'arr�ts, temps double top min, max et moyen
    'Si pas de section de travail,
    'sinon le checksection.value ci-dessus le d�clenche
    If unCheckSection = 0 Then
        frmD.AppelerCheckSectionClick
        'S�lection graphique du rep�re 1
        SelectionnerRepere frmD, 1
    End If
    
    'Affichage de l'onglet histogramme pour bien le retailler
    DoEvents
    frmD.MSChart1.Visible = False
    frmD.TabData.Tab = OngletHistV
    DoEvents
    'Remise en t�te de l'onglet itin�raire de r�f�rence
    frmD.TabData.Tab = OngletItiRef
    frmD.MSChart1.Visible = True
        
    'Mettre � jour les indicateurs de modification
    'pour indiquer que rien n'a �t� modifi� (on vient de l'ouvrir)
    frmD.maModif = False
    frmD.SpreadParcours.ChangeMade = False
    frmD.SpreadRepere.ChangeMade = False
    
    'Mise en gris� ou non des tous les onglets sauf le premier
    'car c'est un nouvel itin�raire que l'on ouvre suivant le nbre de parcours
    'Les onglets vont de 0 � n-1
    For i = 1 To frmD.TabData.Tabs - 1
        frmD.TabData.TabEnabled(i) = (frmD.maColParcours.Count > 0)
    Next i
    
    ' Quitte pour �viter le gestionnaire d'erreur.
    frmMain.MousePointer = vbDefault
    Exit Sub
    
    ' Routine de gestion d'erreur qui �value le num�ro d'erreur.
ErreurLecture:
    
    ' Traite les autres situations ici...
    unMsg = MsgOpenError + unNomFich + Chr(13) + Chr(13) + MsgErreur + Format(Err.Number) + " : " + Err.Description
    If Err.Number = 70 Then unMsg = unMsg + " (" + UCase(MsgDejaOpen) + ")"
    MsgBox unMsg, vbCritical
    'fermeture du fichier et fen�tre
    If (frmD Is Nothing) = False Then
        frmD.maModif = False
        Unload frmD
    End If
    Close #unFichId
    frmMain.MousePointer = vbDefault
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    Exit Sub
End Sub



Public Sub TesterColParcours(uneColParcours)
    Dim unParcours As Parcours, unMsg As String
    Dim uneSep As String, uneDistTot As Long
    
    uneSep = " / "
    For i = 1 To uneColParcours.Count
        unMsg = ""
        Set unParcours = uneColParcours(i)
        unMsg = unParcours.monNom + uneSep + unParcours.monEnqueteur
        unMsg = unMsg + uneSep + Format(unParcours.monNumVeh)
        unMsg = unMsg + uneSep + Format(unParcours.maMeteo)
        unMsg = unMsg + uneSep + unParcours.monJourSemaine
        unMsg = unMsg + uneSep + Format(unParcours.maDate)
        unMsg = unMsg + uneSep + Format(unParcours.monHeureDebut)
        unMsg = unMsg + uneSep + Format(unParcours.monTypeMesure) + Chr(13) + Chr(13)
        unMsg = unMsg + uneSep + Format(unParcours.monCoefEta)
        unMsg = unMsg + uneSep + Format(unParcours.monPasMesure)
        unMsg = unMsg + uneSep + Format(unParcours.maDuree)
        unMsg = unMsg + uneSep + Format(unParcours.maDistPar)
        unMsg = unMsg + uneSep + Format(unParcours.monNbPas)
        unMsg = unMsg + uneSep + Format(unParcours.monFirstPas)
        unMsg = unMsg + uneSep + Format(unParcours.monLastPas)
        unMsg = unMsg + uneSep + Format(UBound(unParcours.monTabAbsRep)) + Chr(13) + Chr(13)
        uneDistTot = 0
        For jj& = 1 To unParcours.monNbPas
            uneDistTot = uneDistTot + unParcours.monTabDist(jj&)
        Next jj&
        unMsg = unMsg + uneSep + "Distance totale = " + Format(uneDistTot)
        unMsg = unMsg + Chr(13) + Chr(13) + "Abs Rep :"
        For j = 1 To UBound(unParcours.monTabAbsRep)
            unMsg = unMsg + uneSep + Format(unParcours.monTabAbsRep(j))
        Next j
        unMsg = unMsg + Chr(13) + Chr(13) + "Tmp Rep :"
        For j = 1 To UBound(unParcours.monTabTempsRep)
            unMsg = unMsg + uneSep + Format(unParcours.monTabTempsRep(j))
        Next j
        If IsEmpty(unParcours.monTabIdRep) = False Then
            'Cas d'un nouveau format de fichier MTB
            unMsg = unMsg + Chr(13) + Chr(13) + "Id Rep :"
            For j = 1 To UBound(unParcours.monTabIdRep)
                unMsg = unMsg + uneSep + Format(unParcours.monTabIdRep(j))
            Next j
        End If
        
        MsgBox unMsg, vbInformation
    Next i
End Sub

Public Sub TesterParcoursMoyen(unParcours As Parcours)
    Dim unMsg As String
    Dim uneSep As String, uneDistTot As Long
    
    uneSep = " / "
    unMsg = ""
    unMsg = unParcours.monNom + uneSep + unParcours.monEnqueteur
    unMsg = unMsg + uneSep + Format(unParcours.monNumVeh)
    unMsg = unMsg + uneSep + Format(unParcours.maMeteo)
    unMsg = unMsg + uneSep + unParcours.monJourSemaine
    unMsg = unMsg + uneSep + Format(unParcours.maDate)
    unMsg = unMsg + uneSep + Format(unParcours.monHeureDebut)
    unMsg = unMsg + uneSep + Format(unParcours.monTypeMesure) + Chr(13) + Chr(13)
    unMsg = unMsg + uneSep + Format(unParcours.monCoefEta)
    unMsg = unMsg + uneSep + Format(unParcours.monPasMesure)
    unMsg = unMsg + uneSep + Format(unParcours.maDuree)
    unMsg = unMsg + uneSep + Format(unParcours.maDistPar)
    unMsg = unMsg + uneSep + Format(unParcours.monNbPas)
    unMsg = unMsg + uneSep + Format(unParcours.monFirstPas)
    unMsg = unMsg + uneSep + Format(unParcours.monLastPas) + Chr(13) + Chr(13)
    uneDistTot = 0
    For j& = 1 To unParcours.monNbPas
        uneDistTot = uneDistTot + unParcours.monTabDist(j&)
    Next j&
    unMsg = unMsg + uneSep + "Distance totale = " + Format(uneDistTot)
    
    MsgBox unMsg, vbInformation
End Sub

Public Function EstNouvelIti(uneForm As Form) As Boolean
    'Fonction retournant vrai si l'itin�raire actif est un nouveau
    'et faux si c'est un itin�raire d�j� existant donc stock� dans un fichier MIT
    If Val(Mid(uneForm.Caption, 12, 1)) > 0 Then
        'Cas d'un nouvel itin�raire ==> Titre de fen�tre = Itin�raire N (N un entier > 0)
        EstNouvelIti = True
    Else
        'Cas d'un itin�raire existant ==> Titre de fen�tre = Itin�raire + nom du fichier
        EstNouvelIti = False
    End If
End Function

Public Function SauverFichier(uneForm As Form, unNomFich As String, unSaveAs As Boolean) As String
    'Sauve l'itin�raire courant dans son fichier .mit si elle existe
    'ou demande un nom de fichier par s�lecteur si c'est un nouvel itin�raire
    
    'Si protection invalide on ne fait rien
   ' If ProtectCheck(2) <> 0 Then
   '     SauverFichier = ""
   '     Exit Function
   ' End If
    
    If EstNouvelIti(uneForm) Or unSaveAs Then
        'Cas d'une nouvelle �tude ou d'un enregistrer sous d'une �tude existante
        unNomFich = frmMain.ChoisirFichier(MsgSaveAs, MsgMitFile, CurDir)
    End If
    
    If unNomFich <> "" Then
        'Cas o� l'utilisateur n'a pas fait annuler
        'dans le s�lecteur de fichiers
        'ou Cas d'une �tude existante (d�j� stock�e dans un fichier .URB)
        '==> unNomFich pas vide
        If EcrireDansFichier(unNomFich, uneForm) Then
            'Mettre � jour liste des fichiers r�cents
            ActualiserListeFichiersRecents unNomFich
            'Mettre � jour les indicateurs de modification
            'pour ne pas demander une sauvegarde lors de la fermeture
            'apr�s un Save ou un SaveAs
            uneForm.maModif = False
            uneForm.SpreadParcours.ChangeMade = False
            uneForm.SpreadRepere.ChangeMade = False
        End If
    End If
    'Valeur de retour
    SauverFichier = unNomFich
    DoEvents
End Function

Public Sub ActualiserListeFichiersRecents(unNomFich As String)
    'Mise � jour de la liste des fichiers r�cents (4 maximum)
    'avec le nom de fichier pass� en param�tre
    'Si ce nom n'est pas dans la liste des fichiers r�cents,
    'il devient num�ro 1, donc passe en t�te et le dernier est supprim�
    'de la liste et les autres d�cal�s de 1
    'S'il est dans la liste, il devient num�ro 1, donc passe en t�te et
    'les autres entre l'ancien 1 et nouveau 1 sont d�cal�s de 1
    
    'Recherche s'il est d�j� pr�sent dans les MRU
    'Dans les mnuFileMRU la chaine est du type "&i Nomfichier"
    For i = 0 To 3
        If frmMain.mnuFileMRU(i).Visible Then unePos = i + 1
        If StrComp(unNomFich, Mid(frmMain.mnuFileMRU(i).Caption, 4), vbTextCompare) = 0 Then
            'Comparaison de texte sans distinguer minuscule et majuscule
            unePos = i
            Exit For
        End If
    Next i
    
    'Cas o� le fichier �tait d�j� dans les MRU files et pas en t�te
    'ou absent (traitement idem que s'il �tait en dernier)
    'D�calage de 1 des MRU files entre les num�ros 0 et unePos-1
    If unePos = 4 Then unePos = 3
    For i = unePos To 1 Step -1
        frmMain.mnuFileMRU(i).Caption = "&" + Format(i + 1) + Mid(frmMain.mnuFileMRU(i - 1).Caption, 3)
        frmMain.mnuFileMRU(i).Visible = True
    Next i
    
    'Mise en t�te du fichier en cours
    frmMain.mnuFileMRU(0).Caption = "&1 " + unNomFich
    frmMain.mnuFileMRU(0).Visible = True
    frmMain.mnuFileBar6.Visible = True
End Sub

Public Function EcrireDansFichier(unNomFich As String, uneForm As Form) As Boolean
    'Ecriture dans le fichier unNomFich du contenu de l'itin�raire uneForm
    'Retour Vrai si tout ok, Faux sinon
    Dim unRep As Repere, unPar As Parcours, uneLigTexte As String
    Dim unNbRepTop As Integer, k As Long, j As Long, unFichId As Byte
    Dim unPasXGrad1 As Single, unPasXGrad2 As Single
    Dim unMaxXreel As Single, unMinXreel As Single
    
    EcrireDansFichier = True
    
    ' Active la routine de gestion d'erreur.
    On Error GoTo ErreurEcriture
    
    ' Fermeture du fichier pour d�lock� et ainsi pouvoir �crire dedans.
    If uneForm.monFichId <> 0 Then
        'Cas d'un Site qui n'est pas Sans Nom (Titre Etude + unNum�ro)
        unFichId = uneForm.monFichId
        Close #unFichId
    End If
        
    'Ouvre le fichier en �criture.
    unFichId = FreeFile(0)
    uneForm.monFichId = unFichId
    Open unNomFich For Output As #unFichId
        
    'Remplissage du fichier � partir des donn�es de l'itin�raire (=uneForm)
    '(cf Format de fichier MiTemps .mit)
    With uneForm
        'Ecriture de l'ent�te des fichiers *.mit
        Write #unFichId, "Fichier " + App.Title
        'Ecriture des libell�s des conditions m�t�o
        uneLigTexte = ""
        For i = 1 To uneForm.maColMeteo.Count - 1
            uneLigTexte = uneLigTexte + uneForm.maColMeteo(i) + ","
        Next i
        'Dernier libell� pas de virgule � la fin
        uneLigTexte = uneLigTexte + uneForm.maColMeteo(i)
        Write #unFichId, uneLigTexte
        'Ecriture des min et max total en distance (parcours complet sans section)
        Write #unFichId, .monMinDtot, .monMaxDtot
        
        'Calcul des max et min Temps r�els pour les stocker
        unMaxXreel = .monMaxT
        unMinXreel = .monMinT
        'Calcul des pas de graduations primaires et secondaires et arrondis
        'de la valeur mini � la graduation secondaire juste inf�rieure
        'et de la valeur maxi � la graduation secondaire juste sup�rieure
        ArrondirXMinXMaxGrad2 unPasXGrad1, unPasXGrad2, unMaxXreel, unMinXreel
        .monMaxT = unMaxXreel
        .monMinT = unMinXreel
        
        'Stockage des max et min Vitesses r�els pour les stocker
        unMaxXreel = .monMaxV
        unMinXreel = .monMinV
        'Calcul des pas de graduations primaires et secondaires et arrondis
        'de la valeur mini � la graduation secondaire juste inf�rieure
        'et de la valeur maxi � la graduation secondaire juste sup�rieure
        ArrondirXMinXMaxGrad2 unPasXGrad1, unPasXGrad2, unMaxXreel, unMinXreel
        .monMaxV = unMaxXreel
        .monMinV = unMinXreel
        
        'Ecriture des min et max  en distance, vitesse et temps
        Write #unFichId, .monMinD, .monMaxD, .monMinV, .monMaxV, .monMinT, .monMaxT
        'Ecriture du nom de l'itin�raire et de sa longueur
        Write #unFichId, .TextNomIti.Text, .TextLongIti.Text
        'Ecriture des donn�es de la section de travail
        Write #unFichId, .CheckSection.Value, .ComboRepDebSec.ListIndex + 1, .ComboRepFinSec.ListIndex + 1
        'Ecriture du nombre de rep�res et de parcours (utile pour la lecture du fichier)
        'sur 4 caract�res pour le nb de parcours car on �crase ces 4 caract�res lors
        'd'un import de parcours dans les fichiers MIT ainsi on ne d�borde jamais
        'sur les autres lignes en faisant un seek ici plus un print en mode ajout
        '(= Append, cf BtnImport_Click de la form frmImportMTB)
        uneLigTexte = Format(.maColRepere.Count) + "," + Format(.maColParcours.Count, "000#")
        Print #unFichId, uneLigTexte
        'Ecriture des donn�es des rep�res
        For i = 1 To .maColRepere.Count
            Set unRep = .maColRepere(i)
            Write #unFichId, unRep.monNomLong, unRep.monNomCourt, unRep.monAbsCurv, unRep.monTypeIcone
        Next i
        'Ecriture des donn�es des parcours, pour chaque parcours :
        'Ligne 1 : Les champs modifiables, Ligne 2 : les non-modifiables
        'stock�s dans le spread parcours dont le nombre de rep�res top�s
        'Ligne 3 : le nombre de pas de mesure,
        '          le pas de mesure en secondes,
        '          le premier et dernier pas de mesure
        'Ligne 4 � nb pas  modulo 13, (13 pas de mesure par ligne)
        'Apr�s les N absCurv des rep�res top�s (10 abscurv par ligne)
        'Puis les N Temps passage des rep�res top�s (10 temps par ligne)
        For i = 1 To .maColParcours.Count
            Set unPar = .maColParcours(i)
            EcrireDonneesParcoursDansFichierMIT unFichId, unPar
        Next i
    End With
    
    'Mise � jour du titre de la fenetre �tude courante
    uneForm.Caption = MsgIti0 + unNomFich
    
    'Fermeture du fichier.
    Close #unFichId
        
    'Ouverture du fichier en lock pour �viter deux ouvertures
    Open unNomFich For Input Lock Read Write As #unFichId
    
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    ' Quitte pour �viter le gestionnaire d'erreur.
    Exit Function
    
    ' Routine de gestion d'erreur qui �value le num�ro d'erreur.
ErreurEcriture:
    
    EcrireDansFichier = False
    ' Traite les autres situations ici...
    unMsg = MsgErreur + Format(Err.Number) + " : " + Err.Description
    MsgBox unMsg, vbCritical
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    'fermeture du fichier
    Close #unFichId
    'Ouverture du fichier en lock pour �viter deux ouvertures
    Open unNomFich For Input Lock Read Write As #unFichId
    Exit Function
End Function


Public Sub EcrireDonneesParcoursDansFichierMIT(unFichId As Byte, unPar As Parcours)
    'Proc�dure �crivant les donn�es d'un parcours dans un fichier itin�raire
    'd'extension MIT
    'Ecriture des donn�es des parcours, pour chaque parcours :
    'Ligne 1 : Les champs modifiables, Ligne 2 : les non-modifiables
    'stock�s dans le spread parcours dont le nombre de rep�res top�s
    'Ligne 3 : le nombre de pas de mesure,
    '          le pas de mesure en secondes,
    '          le premier et dernier pas de mesure
    'Ligne 4 � nb pas  modulo 13, (13 pas de mesure par ligne)
    'Apr�s les N absCurv des rep�res top�s (10 abscurv par ligne)
    'Puis les N Temps passage des rep�res top�s (10 temps par ligne)
    Dim j As Long
    
    unNbRepTop = UBound(unPar.monTabAbsRep)
    Write #unFichId, unPar.monNom, unPar.monIsUtil, unPar.monIsParcoursMoyen, unPar.maCouleur, unPar.monEnqueteur, unPar.monNumVeh, unPar.maMeteo, unPar.maDate, unPar.monJourSemaine, unPar.monHeureDebut
    Write #unFichId, unNbRepTop, unPar.maDistPar, unPar.monTabAbsRep(unNbRepTop), unPar.maDuree, unPar.monCoefEta
    Write #unFichId, unPar.monNbPas, unPar.monPasMesure, unPar.monFirstPas, unPar.monLastPas
    
    'Ecriture des pas de mesure, 13 par ligne maxi, s�par�s par une virgule, sauf le dernier de chaque ligne
    'Chaque pas 5 car maxi, avec la virgule 6 car maxi ==> ligne de 78 = 6*13 car maxi, donc <= 80 lisible sur une page d'�diteur
    unNbParLig = 13
    unQuot = unPar.monNbPas \ unNbParLig
    For j = 1 To unQuot
        uneLigTexte = ""
        For k = 1 To unNbParLig - 1
            uneLigTexte = uneLigTexte + FormatterSingle(CSng(unPar.monTabDist((j - 1) * unNbParLig + k))) + ","
            'uneLigTexte = uneLigTexte + Format(CLng(unPar.monTabDist((j - 1) * unNbParLig + k))) + ","
        Next k
        'uneLigTexte = uneLigTexte + Format(CLng(unPar.monTabDist((j - 1) * unNbParLig + k)))
        uneLigTexte = uneLigTexte + FormatterSingle(CSng(unPar.monTabDist((j - 1) * unNbParLig + k)))
        'ci-dessus k = unNbParLig apr�s sortie du for
        Print #unFichId, uneLigTexte
    Next j
    uneLigTexte = ""
    For j = unQuot * unNbParLig + 1 To unPar.monNbPas
        'uneLigTexte = uneLigTexte + Format(CLng(unPar.monTabDist(j)))
        uneLigTexte = uneLigTexte + FormatterSingle(CSng(unPar.monTabDist(j)))
        If j < unPar.monNbPas Then
            uneLigTexte = uneLigTexte + ","
        End If
    Next j
    If uneLigTexte <> "" Then
        Print #unFichId, uneLigTexte
    End If
    
    'Ecriture des abs curv des rep�res top�s, 10 par ligne maxi, s�par�s par une virgule, sauf le dernier de chaque ligne
    'Chaque abs 7 car maxi, avec la virgule 8 car maxi ==> ligne de 80=10*8 car maxi, donc <= 80 lisible sur une page d'�diteur
    unNbParLig = 10
    unQuot = unNbRepTop \ unNbParLig
    For j = 1 To unQuot
        uneLigTexte = ""
        For k = 1 To unNbParLig - 1
            uneLigTexte = uneLigTexte + Format(CLng(unPar.monTabAbsRep((j - 1) * unNbParLig + k))) + ","
        Next k
        uneLigTexte = uneLigTexte + Format(CLng(unPar.monTabAbsRep((j - 1) * unNbParLig + k)))
        'ci-dessus k = unNbParLig apr�s sortie du for
        Print #unFichId, uneLigTexte
    Next j
    uneLigTexte = ""
    For j = unQuot * unNbParLig + 1 To unNbRepTop
        uneLigTexte = uneLigTexte + Format(CLng(unPar.monTabAbsRep(j)))
        If j < unNbRepTop Then
            uneLigTexte = uneLigTexte + ","
        End If
    Next j
    If uneLigTexte <> "" Then
        Print #unFichId, uneLigTexte
    End If

    'Ecriture des temps de passage des rep�res top�s, 10 par ligne maxi, s�par�s par une virgule, sauf le dernier de chaque ligne
    'Chaque temps 7 car maxi, avec la virgule 8 car maxi ==> ligne de 80=10*8 car maxi, donc <= 80 lisible sur une page d'�diteur
    unNbParLig = 10
    unQuot = unNbRepTop \ unNbParLig
    For j = 1 To unQuot
        uneLigTexte = ""
        For k = 1 To unNbParLig - 1
            'uneLigTexte = uneLigTexte + Format(CLng(unPar.monTabTempsRep((j - 1) * unNbParLig + k))) + ","
            uneLigTexte = uneLigTexte + FormatterSingle(CSng(unPar.monTabTempsRep((j - 1) * unNbParLig + k))) + ","
       Next k
        'uneLigTexte = uneLigTexte + Format(CLng(unPar.monTabTempsRep((j - 1) * unNbParLig + k)))
        uneLigTexte = uneLigTexte + FormatterSingle(CSng(unPar.monTabTempsRep((j - 1) * unNbParLig + k)))
        'ci-dessus k = unNbParLig apr�s sortie du for
        Print #unFichId, uneLigTexte
    Next j
    uneLigTexte = ""
    For j = unQuot * unNbParLig + 1 To unNbRepTop
        'uneLigTexte = uneLigTexte + Format(CLng(unPar.monTabTempsRep(j)))
        uneLigTexte = uneLigTexte + FormatterSingle(CSng(unPar.monTabTempsRep(j)))
        If j < unNbRepTop Then
            uneLigTexte = uneLigTexte + ","
        End If
    Next j
    If uneLigTexte <> "" Then
        Print #unFichId, uneLigTexte
    End If
End Sub

Public Function RecupererContenuEtude(unNomFich As String, uneColRep As ColRepere, uneColPar As ColParcours) As Integer
    'Ouvre l'�tude contenue dans le fichier pass� en param�tre
    'pour r�cup�rer les rep�res et parcours dans les deux collections
    'pass�es en param�tres
    'On retourne la ligne o� se trouve les nombres de rep�res et de parcours
    'si retour = 0 la lecture de l'itin�raire ou �tude a �chou�
    
    Dim unRep As Repere, unBool As Boolean, unBool2 As Boolean
    Dim unMinD As Single, unMaxD As Single, unNumLigNbPar As Integer
    Dim unPar As Parcours, unParMoyen As Parcours
    Dim uneAbsCurv As Long, unTypeIco As Byte, unY1 As Long, unY2 As Long
    Dim unNbPas As Long, unNbRepTop As Long, j As Long
    Dim unNbRep As Integer, unNbPar As Integer
    Dim unLong1 As Long, unLong2 As Long, unLong3 As Long, unByte As Byte
    Dim unInt1 As Integer, unInt2 As Integer, unInt3 As Integer
    Dim uneString As String, uneString1 As String, uneString2 As String
    Dim uneString3 As String, uneHeure As Date, uneDate As Date
    Dim frmD As frmDocument, unePos As Integer
    Dim unTabAbsRep As Variant, unTabTempsRep As Variant
    Dim unTabSng(1 To NbPasMax) As Single 'Pour stocker les lectures de single
    Dim unCheckSection As Integer, unIndRepDeb As Integer, unIndRepFin As Integer
    
    'Initialisation de la valeur de retour � 0
    'si on va au bout elle vaudra la ligne contenant les nombres
    'de rep�res et de parcours, sinon 0
    RecupererContenuEtude = 0
    
    'Si protection invalide on ne fait rien
    'If ProtectCheck(2) <> 0 Then Exit Function
    
    'Lecture du fichier .mit
    ' Active la routine de gestion d'erreur.
    'MsgBox "Suppression du On Error GoTo ErreurLecture2"
    On Error GoTo ErreurLecture2
    
    'Ouverture du fichier en lecture lock�e pour �viter deux ouvertures
    unFichId = FreeFile(0)
    Open unNomFich For Input Lock Read Write As #unFichId
       
    'Lecture de l'ent�te des fichiers *.mit
    Input #unFichId, uneString
    If uneString <> "Fichier " + App.Title Then
        'Cas d'un fichier qui n'est pas un fichier MiTemps
        '===> Fermeture du fichier.
        Close #unFichId
        MsgBox MsgErreur + MsgFileNotFile + App.Title + " version " + Format(App.Major) + "." + Format(App.Minor), vbCritical
        ' D�sactive la r�cup�ration d'erreur.
        On Error GoTo 0
        Exit Function
    Else
        'Cas d'un fichier MiTemps *.Mit de la version 3.0
        '1�re ligne du fichier MIT = "Fichier MiTemps"
        
        'R�cup�ration des libell�s des conditions m�t�o
        Input #unFichId, uneString
        For i = 0 To 7
            unePos = InStr(1, uneString, ",")
            uneString = Mid(uneString, unePos + 1)
        Next i
        
        'R�cup�ration des min et max total en distance (parcours complet sans section)
        Input #unFichId, unTabSng(1), unTabSng(2)
        'R�cup�ration des min et max en distance, vitesse et temps
        Input #unFichId, unMinD, unMaxD, unTabSng(3), unTabSng(4), unTabSng(5), unTabSng(6)
        
        'R�cup�ration du nom de l'itin�raire et de sa longueur
        Input #unFichId, uneString, uneString1
        
        'R�cup�ration des donn�es de la section de travail
        Input #unFichId, unCheckSection, unIndRepDeb, unIndRepFin
        
        'Stockage de la position de lecture juste avant la lecture du nombre
        'de rep�res et parcours, on s'en sert lors de l'importation pour �crire
        'dans le fichier mit de l'itin�raire charg� le nombre de parcours + 1
        unNumLigNbPar = Seek(unFichId)
        
        'R�cup�ration du nombre de rep�res et de parcours
        Input #unFichId, unNbRep, unNbPar
        'R�cup�ration, stockage et cr�ation des rep�res
        For i = 1 To unNbRep
            Input #unFichId, uneString, uneString1, uneAbsCurv, unTypeIco
            Set unRep = uneColRep.Add(uneString, uneString1, uneAbsCurv, unTypeIco)
        Next i
        'Remise � z�ro de unNbRep
        unNbRep = 0
        
        'R�cup�ration des donn�es des parcours, pour chaque parcours :
        'Ligne 1 : Les champs modifiables, Ligne 2 : les non-modifiables
        'stock�s dans le spread parcours dont le nombre de rep�res top�s
        'Ligne 3 : le nombre de pas de mesure,
        '          le pas de mesure en secondes,
        '          le premier et dernier pas de mesure
        'Ligne 4 � nb pas  modulo 13, (13 pas de mesure par ligne)
        'Apr�s les N absCurv des rep�res top�s (10 abscurv par ligne)
        'Puis les N Temps passage des rep�res top�s (10 temps par ligne)
        For i = 1 To unNbPar
            Input #unFichId, uneString, unBool, unBool2, unLong1, uneString1, uneString2, unByte, uneDate, uneString3, uneHeure
            Set unPar = uneColPar.Add(uneString, unLong1)
            unPar.monIsUtil = unBool
            unPar.monIsParcoursMoyen = unBool2
            unPar.monEnqueteur = uneString1
            unPar.monNumVeh = uneString2
            unPar.maMeteo = unByte
            unPar.maDate = uneDate
            unPar.monJourSemaine = uneString3
            unPar.monHeureDebut = uneHeure
            
            Input #unFichId, unNbRepTop, unLong1, unLong2, unLong3, unTabSng(1)
            unPar.maDistPar = unLong1
            unPar.maDuree = unLong3
            unPar.monCoefEta = unTabSng(1)
            
            'Stockage dans unNbRep du nombre de rep�res maxis
            If unNbRepTop > unNbRep Then unNbRep = unNbRepTop
            
            Input #unFichId, unNbPas, unLong1, unInt1, unInt2
            unPar.monNbPas = unNbPas
            unPar.monPasMesure = unLong1
            unPar.monFirstPas = unInt1
            unPar.monLastPas = unInt2
                        
            'R�cup�ration des pas de mesure
            For j = 1 To unNbPas
                'Lecture de la donn�e
                Input #unFichId, unTabSng(1)
                unPar.monTabDist(j) = unTabSng(1)
            Next j
                        
            'Allocation dynamique des tableaux li�s aux rep�res top�s
            unTabAbsRep = unPar.monTabAbsRep
            unTabTempsRep = unPar.monTabTempsRep
            ReDim unTabAbsRep(1 To unNbRepTop)
            ReDim unTabTempsRep(1 To unNbRepTop)
            
            'R�cup�ration des abs curv des rep�res top�s
            For j = 1 To unNbRepTop
                'Lecture de la donn�e
                Input #unFichId, unLong1
                unTabAbsRep(j) = unLong1
            Next j
        
            'R�cup�ration des temps de passage des rep�res top�s
            For j = 1 To unNbRepTop
                'Lecture de la donn�e
                Input #unFichId, unTabSng(1)
                unTabTempsRep(j) = unTabSng(1)
            Next j
            
            'Affectation des pointeurs sur les tableaux du parcours
            unPar.monTabAbsRep = unTabAbsRep
            unPar.monTabTempsRep = unTabTempsRep
        Next i
    End If
        
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
        
    'Vidage de tous les events en attente
    DoEvents
        
    ' Quitte pour �viter le gestionnaire d'erreur en retournant que tout
    's'est bien pass�, donc le num�ro de la ligne o� on lit le nombre de parcours
    Close #unFichId
    RecupererContenuEtude = unNumLigNbPar
    Exit Function
    
    ' Routine de gestion d'erreur qui �value le num�ro d'erreur.
ErreurLecture2:
    
    ' Traite les autres situations ici...
    unMsg = MsgOpenError + unNomFich + Chr(13) + Chr(13) + MsgErreur + Format(Err.Number) + " : " + Err.Description
    If Err.Number = 70 Then unMsg = unMsg + " (" + UCase(MsgDejaOpen) + ")"
    MsgBox unMsg, vbCritical
    'Fermeture du fichier et retour 0 de la fonction car �chec lors de la lecture
    RecupererContenuEtude = 0
    Close #unFichId
    ' D�sactive la r�cup�ration d'erreur.
    On Error GoTo 0
    Exit Function
End Function

