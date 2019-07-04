Attribute VB_Name = "ModuleFichier"
Public Function LireFichierMTB(unNomFich As String) As Boolean
    'Lecture du fichier MTB passé en paramètre
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
    'Effectue la boucle jusqu'à la fin du fichier.
    unNbData = 0
    unNbParLu = 1
    Do While Not EOF(unFichId)
        ' Lit les données dans une variable.
        Input #unFichId, uneString
        'Incrémentation du nombre de données pour ce parcours
        unNbData = unNbData + 1
        unePos = InStr(1, uneString, "!", vbTextCompare)
        If unePos > 0 Then
            'Cas du premier ! trouvé (code pour lire les vieux fichiers MTB contenant des !!)
            Input #unFichId, uneStrTmp 'passage au ! suivant
            Input #unFichId, uneStrTmp 'Lecture du reste de la donnée coupé par les !!
            'Reconstruction de la valeur globale
            If uneString = "!" Then
                uneString = uneStrTmp
            Else
                uneString = Mid(uneString, 1, Len(uneString) - 1) + uneStrTmp
            End If
        End If

        'Remplissage des données du parcours
        Select Case unNbData  ' Évalue unNbData.
            Case 1
                If uneString <> "" Then
                    'Création du parcours issu du MTB
                    'avec affectation du nom et d'une couleur par défaut
                    'Affectation d'une couleur par défaut, on commence à 9
                    'pour éviter le gris (cf aide sur fonction QBColor)
                    Set unParcours = maColParcoursMTB.Add(uneString, QBColor(9 + unNbParLu Mod 6))
                    'Mise à zéro du pas stockant le début d'un arrêt
                    unNumPasDebArret = 0
                End If
                'Si le fichier MTB a un 0D0A en fin de fichier, la dernière lecture
                'donne une chiane vide pour le nom, c'est ainsi que l'on repère
                'ces fichiers MTB qui ont été sauvé par un éditeur DOS ou Windows
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
                'On enlève les centièmes de secondes de l'heure de départ
                'Date VB va jusqu'au seconde et le spread aussi
                'On supprime les trois derniers caractères hh:MM:ss:mm
                'donc :mm
                unParcours.monHeureDebut = TimeValue(Mid(uneString, 1, Len(uneString) - 3))
            Case 8
                'Lecture du type de mesure (1 car = D presque toujours)
                unParcours.monTypeMesure = uneString
            Case 9
                '9 (= lecture du nbre de repères théoriques) on ne fait rien
                '==> on passe à la lecture suivante
                uneString = uneString
            Case 10
                'On remplace le caractère décimale en cours par le point
                'sinon CSng plante, en effet le séparateur est le point
                'dans les fichiers MTB
                unePosSepDec = InStr(1, uneString, ".")
                'Récup de la partie entière à gauche du point décimale
                unePartEnt = Format(Mid(uneString, 1, unePosSepDec - 1))
                'Récup de la partie décimale (4 chiffres maxi) à gauche du point décimale
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
                'On divise car les données sont déjà multipliés par le coef d'étalonnage
                'Ainsi les distances stockées sont indépendants du coef d'étalonnage
            Case 14
                'Stockage du nombre de pas de mesure
                unNbPas = CLng(uneString)
                unParcours.monNbPas = unNbPas
                If unNbPas > NbPasMax Then
                    'Cas ou le nombre de mesures dépasse le nb maxi
                    'fixé, on sort sans rien faire
                    MsgBox "Le nombre de mesures, valant " + Format(unNbPas) + ", du parcours " + Format(unParcours.monNom) + " dépasse le nombre de mesures maximun fixé à " + Format(NbPasMax), vbCritical
                    LireFichierMTB = False
                    Exit Function
                End If
            Case 15
                unParcours.monFirstPas = CInt(uneString)
            Case 16
                unParcours.monLastPas = CInt(uneString)
            Case 17
                'Stockage du nombre de repères topés
                unNbRepTop = CInt(uneString)
                If unNbRepTop < 2 Then
                    'Obligation d'avoir au moins deux repères
                    MsgBox "Le nombre de repères topés du parcours lu en position " + Format(unNbParLu) + " devrait être supérieur ou égal à 2", vbInformation
                    'on se sort plus car cela ne géne en rien le reste et en plus
                    'on peut continuer de lire les parcours suivants
                    'LireFichierMTB = False
                    'Exit Function
                End If
                'Allocation dynamique des tableaux liés aux repères topés
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
                    'Test de Fin de lecture des données de ce parcours
                    'Lecture du temps de passage du dernier repère
                    unTabTempsRep(unNbRepTop) = CLng(uneString)
                    
                    If unFormatMTB = NewMTB Then
                        'Cas du nouveau format de fichier
                        'On lit les champs d'identification des repères
                        For i = 1 To unNbRepTop
                            Input #unFichId, uneString
                            unTabIdRep(i) = uneString
                        Next i
                        'Affectation des pointeurs sur le tableau
                        'des champs d'identification des repères du parcours
                        unParcours.monTabIdRep = unTabIdRep
                    End If
                    
                    'On remet unNbData à 0 pour la lecture des données
                    'des parcours suivants éventuelles
                    unNbData = 0
                
                    'Incrémentation du nombre de parcours
                    unNbParLu = unNbParLu + 1
                    
                    'Affectation des pointeurs sur les tableaux du parcours
                    unParcours.monTabTempsRep = unTabTempsRep
                    unParcours.monTabAbsRep = unTabAbsRep
                ElseIf unNbData > 17 + unNbPas + unNbRepTop Then
                    'Lecture des temps de passages des repères.
                    unTabTempsRep(unNbData - 17 - unNbPas - unNbRepTop) = CLng(uneString)
                ElseIf unNbData > 17 + unNbPas Then
                    'Lecture des abscisses des repères.
                    unTabAbsRep(unNbData - 17 - unNbPas) = CLng(uneString) / unParcours.monCoefEta
                    'On divise car les données sont déjà multipliés par le coef d'étalonnage
                    'Ainsi les abscisses stockées sont indépendants du coef d'étalonnage
                ElseIf unNbData > 17 Then
                    'Lecture des distances parcourues par pas
                    unParcours.monTabDist(unNbData - 17) = CLng(uneString) / unParcours.monCoefEta
                    'On divise car les données sont déjà multipliés par le coef d'étalonnage
                    'Ainsi les abscisses stockées sont indépendants du coef d'étalonnage
                    
                    'Calcul du nombre et du temps d'arrêts
                    If CLng(uneString) = 0 Then
                        'Cas où on est arrêté
                        If unNumPasDebArret = 0 Then
                            'Stockage du pas débutant l'arrêt, tjs > 0
                            unNumPasDebArret = unNbData - 17
                            unParcours.monNbArret = unParcours.monNbArret + 1
                        End If
                        'Cumul du temps d'arrêt en dixième de seconde,
                        'le pas de mesure est en seconde
                        If unNbData - 17 = 1 Then
                            unParcours.monTpsArret = unParcours.monTpsArret + unParcours.monFirstPas
                        ElseIf unNbData - 17 = unParcours.monNbPas Then
                            unParcours.monTpsArret = unParcours.monTpsArret + unParcours.monLastPas
                        Else
                            unParcours.monTpsArret = unParcours.monTpsArret + unParcours.monPasMesure * 10
                        End If
                    Else
                        'Cas où on n'est pas arrêté
                        'Remise à zéro du pas stockant le début de l'arrêt
                        unNumPasDebArret = 0
                    End If
                End If
        End Select
    Loop
    
    'Ferme le fichier
    Close #unFichId
    
    'Désactive la récupération d'erreur.
    On Error GoTo 0
    'Sortie de la procédure pour éviter le passage
    'dans la gestion d'erreur
    Exit Function

ErreurLireMTB:
    'Cas où erreur de lecture du fichier MTB
    If unNbData = 0 And unNbParLu > maColParcoursMTB.Count Then
        'Lecture en fin de fichier normal, ce n'est pas une erreur
        'les mtb finissent tous par une virgule, on a lu le dernier champ
        'le nbdata remis à zéro mais le nbParLu incrémente avant le add dans
        'la collection des parcours, donc c'est le cas où on lit après la fin
        'du fichier, apès la dernière virgule
        LireFichierMTB = True
    Else
        'Cas des erreurs
        MsgBox MsgErreur + Format(Err.Number) + " " + MsgIn + "LireFichierMTB / " + Err.Description, vbCritical
        ViderColParcours maColParcoursMTB
        LireFichierMTB = False
    End If
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    'Ferme le fichier
    Close #unFichId
    Exit Function
End Function

Public Function IsNewFichierMTB(unNomFich As String) As Byte
    'Function indiquant si le fichier MTB passé en paramètre
    'est au nouveau format de fichier ou pas
    'Le nouveau format comporte des champs d'identification de repère
    '(1 caractère par repère) en plus
    'Valeur de retour :
    '   OldMTB si vieux format
    '   NewMTB si nouveau format
    '   BadMTB si mauvais format de fichier ( tous les autres cas)
    'Ces constantes sont définies dans le module ModuleMain
    
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
    'Effectue la boucle jusqu'à la fin du fichier.
    unNbData = 0
    uneFinLecture = False
    Do While Not uneFinLecture
        ' Lit les données dans une variable.
        Input #unFichId, uneString
        unNbData = unNbData + 1
        unePos = InStr(1, uneString, "!", vbTextCompare)
        If unePos > 0 Then
            'Cas du premier ! trouvé (code pour lire les vieux fichiers MTB contenant des !!)
            Input #unFichId, uneStrTmp 'passage au ! suivant
            Input #unFichId, uneStrTmp 'Lecture du reste de la donnée coupé par les !!
            'Reconstruction de la valeur globale
            If uneString = "!" Then
                uneString = uneStrTmp
            Else
                uneString = Mid(uneString, 1, Len(uneString) - 1) + uneStrTmp
            End If
        End If

        'Récupération de certaines données du parcours
        If unNbData = 14 Then
            'Stockage du nombre de pas de mesure
            unNbPas = CLng(uneString)
        ElseIf unNbData = 17 Then
            'Stockage du nombre de repères topés
            unNbRepTop = CInt(uneString)
        ElseIf unNbData = 17 + unNbPas + unNbRepTop * 2 Then
            'Test de Fin de lecture des données de ce parcours
            'Lecture des unNbRepTop champs suivants pour voir si c'est
            'un nouveau de fichier avec les identifications de repères
            'sur un caractère.
            'Si ces unNbRepTop champs ont tous une longueur <= 1
            'alors nouveau format de fichier sinon ancien fichier
            'Si ancien fichier et pas d'autres parcours,
            '==> le champ lu suivant déclenche EOF
            For i = 1 To unNbRepTop
                Input #unFichId, uneString
                'Si on dépasse la fin du fichier ==> Vieux format fichier
                'on part dans la gestion d'erreur ErreurNewMTB plus bas
                'Si la longueur chaine est > 1 ==> Vieux format de fichier,
                'pas de champ identification des repères
                If Len(uneString) > 1 Then
                    IsNewFichierMTB = OldMTB
                    Exit For
                End If
            Next i
            
            'Arrêt de la boucle de lecture
            uneFinLecture = True
        End If
    Loop
    
    'Ferme le fichier
    Close #unFichId
    
    'Désactive la récupération d'erreur.
    On Error GoTo 0
    'Sortie de la procédure pour éviter le passage
    'dans la gestion d'erreur
    Exit Function

ErreurNewMTB:
    'Cas où erreur de lecture du fichier MTB
    If unNbData = 17 + unNbPas + unNbRepTop * 2 Then
        'Lecture en fin de fichier normal, ce n'est pas une erreur
        'c'est un fichier MTB au vieux format
        IsNewFichierMTB = OldMTB
    Else
        'Cas des autres erreurs
        MsgBox MsgErreur + Format(Err.Number) + " " + MsgIn + "IsNewFichierMTB / " + Err.Description, vbCritical
        IsNewFichierMTB = BadMTB
    End If
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    'Ferme le fichier
    Close #unFichId
    Exit Function
End Function

Public Sub OuvrirEtude(unNomFich As String)
    'Ouvre l'étude contenue dans le fichier passé en paramètre
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
    
    'Si protection invalide on ne fait rien ancien système neutralisé
    'If ProtectCheck(2) <> 0 Then Exit Sub
    
    'Lecture du fichier .mit
    ' Active la routine de gestion d'erreur.
    'MsgBox "Suppression du On Error GoTo ErreurLecture"
    On Error GoTo ErreurLecture
    
    'Ouverture du fichier en lecture lockée pour éviter deux ouvertures
    unFichId = FreeFile(0)
    Open unNomFich For Input Lock Read Write As #unFichId
    
    'Création de la nouvelle fenêtre qui contiendra l'étude
    Set frmD = New frmDocument
    frmD.Visible = False 'pour corriger un bug bizarre VB
    'Titre de la fenêtre itinéraire
    frmD.Caption = MsgIti0 + unNomFich
    
    'Stockage du fichier id
    frmD.monFichId = unFichId
    
    'Lecture de l'entête des fichiers *.mit
    Input #unFichId, uneString
    If uneString <> "Fichier " + App.Title Then
        'Cas d'un fichier qui n'est pas un fichier MiTemps
        '===> Fermeture du fichier.
        Close #unFichId
        MsgBox MsgErreur + MsgFileNotFile + App.Title + " version " + Format(App.Major) + "." + Format(App.Minor), vbCritical
        ' Désactive la récupération d'erreur.
        On Error GoTo 0
        Exit Sub
    Else
        'Cas d'un fichier MiTemps *.Mit de la version 3.0
        '1ère ligne du fichier MIT = "Fichier MiTemps"
        
        'Récupération des libellés des conditions météo
        Input #unFichId, uneString
        For i = 0 To 7
            unePos = InStr(1, uneString, ",")
            frmD.maColMeteo.Add Mid(uneString, 1, unePos - 1)
            uneString = Mid(uneString, unePos + 1)
        Next i
        'Dernier libellé
        frmD.maColMeteo.Add uneString
        
        'Récupération des min et max total en distance (parcours complet sans section)
        Input #unFichId, unTabSng(1), unTabSng(2)
        frmD.monMinDtot = unTabSng(1)
        frmD.monMaxDtot = unTabSng(2)
        'Récupération des min et max en distance, vitesse et temps
        'Input #unFichId, unTabSng(1), unTabSng(2), unTabSng(3), unTabSng(4), unTabSng(5), unTabSng(6)
        Input #unFichId, unMinD, unMaxD, unTabSng(3), unTabSng(4), unTabSng(5), unTabSng(6)
        frmD.monMinD = unMinD 'unTabSng(1)
        frmD.monMaxD = unMaxD 'unTabSng(2)
        frmD.monMinV = unTabSng(3)
        frmD.monMaxV = unTabSng(4)
        frmD.monMinT = unTabSng(5)
        frmD.monMaxT = unTabSng(6)
        
        'Récupération du nom de l'itinéraire et de sa longueur
        Input #unFichId, uneString, uneString1
        frmD.TextNomIti.Text = uneString
        frmD.TextLongIti.Text = uneString1
        frmD.maLongIti = CLng(uneString1)
        
        'Récupération des données de la section de travail
        Input #unFichId, unCheckSection, unIndRepDeb, unIndRepFin
        
        'On remet à zéro le tableau de repères
        frmD.SpreadRepere.MaxRows = 0
        'Récupération du nombre de repères et de parcours
        Input #unFichId, unNbRep, unNbPar
        'Récupération et création des repères
        For i = 1 To unNbRep
            Input #unFichId, uneString, uneString1, uneAbsCurv, unTypeIco
            Set unRep = CreerRepere(frmD, uneString, uneString1, uneAbsCurv, unTypeIco)
        Next i
        'Remise à zéro de unNbRep
        unNbRep = 0
        
        'Récupération des données des parcours, pour chaque parcours :
        'Ligne 1 : Les champs modifiables, Ligne 2 : les non-modifiables
        'stockés dans le spread parcours dont le nombre de repères topés
        'Ligne 3 : le nombre de pas de mesure,
        '          le pas de mesure en secondes,
        '          le premier et dernier pas de mesure
        'Ligne 4 à nb pas  modulo 13, (13 pas de mesure par ligne)
        'Aprés les N absCurv des repères topés (10 abscurv par ligne)
        'Puis les N Temps passage des repères topés (10 temps par ligne)
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
            
            'Stockage dans unNbRep du nombre de repères maxis
            If unNbRepTop > unNbRep Then unNbRep = unNbRepTop
            
            Input #unFichId, unNbPas, unLong1, unInt1, unInt2
            unPar.monNbPas = unNbPas
            unPar.monPasMesure = unLong1
            unPar.monFirstPas = unInt1
            unPar.monLastPas = unInt2
                        
            'Récupération des pas de mesure
            For j = 1 To unNbPas
                'Lecture de la donnée
                Input #unFichId, unTabSng(1)
                unPar.monTabDist(j) = unTabSng(1)
            Next j
                        
            'Allocation dynamique des tableaux liés aux repères topés
            unTabAbsRep = unPar.monTabAbsRep
            unTabTempsRep = unPar.monTabTempsRep
            ReDim unTabAbsRep(1 To unNbRepTop)
            ReDim unTabTempsRep(1 To unNbRepTop)
            
            'Récupération des abs curv des repères topés
            For j = 1 To unNbRepTop
                'Lecture de la donnée
                Input #unFichId, unLong1
                unTabAbsRep(j) = unLong1
            Next j
        
            'Récupération des temps de passage des repères topés
            For j = 1 To unNbRepTop
                'Lecture de la donnée
                Input #unFichId, unTabSng(1)
                unTabTempsRep(j) = unTabSng(1)
            Next j
            
            'Affectation des pointeurs sur les tableaux du parcours
            unPar.monTabAbsRep = unTabAbsRep
            unPar.monTabTempsRep = unTabTempsRep
        Next i
    
        If unNbPar > 0 Then
            If frmD.maColParcours(1).monIsParcoursMoyen = False Then
                'Cas où le premier parcours n'est pas le parcours moyen
                '==> il faut le créer et le calculer
                'Cela se produit uniquement si on importe des parcours dans
                'un fichier itinéraire n'ayant pas de parcours ou n'ayant jamais
                'été ouvert puis sauvegarder (car le parcours moyen est alors créé)
                
                'Création du parcours moyen et
                'Ajout en tête dans les parcours du nouvel itinéraire
                'Le parcours moyen sera toujours celui en première position
                ' qbcolor(0) = noir
                Set unPar = New Parcours
                frmD.maColParcours.Ajouter unPar, True
                'Indication que ce parcours créé est le parcours moyen
                'et qu'il sera utilisé
                unPar.monIsParcoursMoyen = True
                unPar.monIsUtil = True
            Else
                Set unPar = frmD.maColParcours(1)
            End If
            
            'Allocation dynamique des tableaux liés aux repères topés
            'Le parcours moyen aura un nombre top valant le plus grand nombre
            'de top de tous les autres parcours
            unTabAbsRep = unPar.monTabAbsRep
            unTabTempsRep = unPar.monTabTempsRep
            ReDim unTabAbsRep(1 To unNbRep)
            ReDim unTabTempsRep(1 To unNbRep)
            'Affectation des pointeurs sur les tableaux du parcours
            unPar.monTabAbsRep = unTabAbsRep
            unPar.monTabTempsRep = unTabTempsRep
            
            'Calcul des caractéristiques générales
            'et du tableau de distances parcourues par pas du parcours
            If unCheckSection = 0 Then
                'Stockage des abs début et fin du parcours
                unY1 = -100
                unY2 = 1000000
            Else
                'Stockage des abs début et fin de la section de travail du parcours
                unY1 = frmD.maColRepere(unIndRepDeb).monAbsCurv
                unY2 = frmD.maColRepere(unIndRepFin).monAbsCurv
            End If
            ActualiserParcoursMoyen frmD.maColParcours(1), frmD.maColParcours, unY1, unY2
        End If
        
        'Remplissage du spread parcours de la fenêtre itinéraire
        RemplirSpreadParcours frmD
        'Remplissage du spread parcours avec les libellés météo
        'de la fenêtre itinéraire
        RemplirMeteoSpreadParcours frmD
    End If
        
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
        
    'Mettre à jour liste des fichiers récents
    ActualiserListeFichiersRecents unNomFich
    
    'Vidage de tous les events en attente
    DoEvents
    'Mise en zoom total pour éviter les redessins successifs lors de la
    'création des repères
    frmD.monMinD = frmD.monMinDtot
    frmD.monMaxD = frmD.monMaxDtot
    'Affichage de la fenêtre qui déclenche le dessin des repères
    frmD.Show
    'Remise de l'englobant stocké lors de la sauvegarde du fichier mit
    frmD.monMinD = unMinD
    frmD.monMaxD = unMaxD
        
    'Mise à jour des données de la section de travail car c'est CreerRepere
    'qui crée les éléments des combobox début et fin de section
    frmD.ComboRepDebSec.ListIndex = unIndRepDeb - 1
    frmD.ComboRepFinSec.ListIndex = unIndRepFin - 1
    'Stockage dans le tag des dernières valeurs valides
    frmD.ComboRepDebSec.Tag = frmD.ComboRepDebSec.Text
    frmD.ComboRepFinSec.Tag = frmD.ComboRepFinSec.Text
    frmD.CheckSection.Value = unCheckSection
    
    'Appeler du click event de CheckSection pour recalcul l'englobant et
    'les vitesses, temps d'arrêts, temps double top min, max et moyen
    'Si pas de section de travail,
    'sinon le checksection.value ci-dessus le déclenche
    If unCheckSection = 0 Then
        frmD.AppelerCheckSectionClick
        'Sélection graphique du repère 1
        SelectionnerRepere frmD, 1
    End If
    
    'Affichage de l'onglet histogramme pour bien le retailler
    DoEvents
    frmD.MSChart1.Visible = False
    frmD.TabData.Tab = OngletHistV
    DoEvents
    'Remise en tête de l'onglet itinéraire de référence
    frmD.TabData.Tab = OngletItiRef
    frmD.MSChart1.Visible = True
        
    'Mettre à jour les indicateurs de modification
    'pour indiquer que rien n'a été modifié (on vient de l'ouvrir)
    frmD.maModif = False
    frmD.SpreadParcours.ChangeMade = False
    frmD.SpreadRepere.ChangeMade = False
    
    'Mise en grisé ou non des tous les onglets sauf le premier
    'car c'est un nouvel itinéraire que l'on ouvre suivant le nbre de parcours
    'Les onglets vont de 0 à n-1
    For i = 1 To frmD.TabData.Tabs - 1
        frmD.TabData.TabEnabled(i) = (frmD.maColParcours.Count > 0)
    Next i
    
    ' Quitte pour éviter le gestionnaire d'erreur.
    frmMain.MousePointer = vbDefault
    Exit Sub
    
    ' Routine de gestion d'erreur qui évalue le numéro d'erreur.
ErreurLecture:
    
    ' Traite les autres situations ici...
    unMsg = MsgOpenError + unNomFich + Chr(13) + Chr(13) + MsgErreur + Format(Err.Number) + " : " + Err.Description
    If Err.Number = 70 Then unMsg = unMsg + " (" + UCase(MsgDejaOpen) + ")"
    MsgBox unMsg, vbCritical
    'fermeture du fichier et fenêtre
    If (frmD Is Nothing) = False Then
        frmD.maModif = False
        Unload frmD
    End If
    Close #unFichId
    frmMain.MousePointer = vbDefault
    ' Désactive la récupération d'erreur.
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
    'Fonction retournant vrai si l'itinéraire actif est un nouveau
    'et faux si c'est un itinéraire dèjà existant donc stocké dans un fichier MIT
    If Val(Mid(uneForm.Caption, 12, 1)) > 0 Then
        'Cas d'un nouvel itinéraire ==> Titre de fenêtre = Itinéraire N (N un entier > 0)
        EstNouvelIti = True
    Else
        'Cas d'un itinéraire existant ==> Titre de fenêtre = Itinéraire + nom du fichier
        EstNouvelIti = False
    End If
End Function

Public Function SauverFichier(uneForm As Form, unNomFich As String, unSaveAs As Boolean) As String
    'Sauve l'itinéraire courant dans son fichier .mit si elle existe
    'ou demande un nom de fichier par sélecteur si c'est un nouvel itinéraire
    
    'Si protection invalide on ne fait rien
   ' If ProtectCheck(2) <> 0 Then
   '     SauverFichier = ""
   '     Exit Function
   ' End If
    
    If EstNouvelIti(uneForm) Or unSaveAs Then
        'Cas d'une nouvelle étude ou d'un enregistrer sous d'une étude existante
        unNomFich = frmMain.ChoisirFichier(MsgSaveAs, MsgMitFile, CurDir)
    End If
    
    If unNomFich <> "" Then
        'Cas où l'utilisateur n'a pas fait annuler
        'dans le sélecteur de fichiers
        'ou Cas d'une étude existante (déjà stockée dans un fichier .URB)
        '==> unNomFich pas vide
        If EcrireDansFichier(unNomFich, uneForm) Then
            'Mettre à jour liste des fichiers récents
            ActualiserListeFichiersRecents unNomFich
            'Mettre à jour les indicateurs de modification
            'pour ne pas demander une sauvegarde lors de la fermeture
            'après un Save ou un SaveAs
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
    'Mise à jour de la liste des fichiers récents (4 maximum)
    'avec le nom de fichier passé en paramètre
    'Si ce nom n'est pas dans la liste des fichiers récents,
    'il devient numéro 1, donc passe en tête et le dernier est supprimé
    'de la liste et les autres décalés de 1
    'S'il est dans la liste, il devient numéro 1, donc passe en tête et
    'les autres entre l'ancien 1 et nouveau 1 sont décalés de 1
    
    'Recherche s'il est déjà présent dans les MRU
    'Dans les mnuFileMRU la chaine est du type "&i Nomfichier"
    For i = 0 To 3
        If frmMain.mnuFileMRU(i).Visible Then unePos = i + 1
        If StrComp(unNomFich, Mid(frmMain.mnuFileMRU(i).Caption, 4), vbTextCompare) = 0 Then
            'Comparaison de texte sans distinguer minuscule et majuscule
            unePos = i
            Exit For
        End If
    Next i
    
    'Cas où le fichier était dèjà dans les MRU files et pas en tête
    'ou absent (traitement idem que s'il était en dernier)
    'Décalage de 1 des MRU files entre les numéros 0 et unePos-1
    If unePos = 4 Then unePos = 3
    For i = unePos To 1 Step -1
        frmMain.mnuFileMRU(i).Caption = "&" + Format(i + 1) + Mid(frmMain.mnuFileMRU(i - 1).Caption, 3)
        frmMain.mnuFileMRU(i).Visible = True
    Next i
    
    'Mise en tête du fichier en cours
    frmMain.mnuFileMRU(0).Caption = "&1 " + unNomFich
    frmMain.mnuFileMRU(0).Visible = True
    frmMain.mnuFileBar6.Visible = True
End Sub

Public Function EcrireDansFichier(unNomFich As String, uneForm As Form) As Boolean
    'Ecriture dans le fichier unNomFich du contenu de l'itinéraire uneForm
    'Retour Vrai si tout ok, Faux sinon
    Dim unRep As Repere, unPar As Parcours, uneLigTexte As String
    Dim unNbRepTop As Integer, k As Long, j As Long, unFichId As Byte
    Dim unPasXGrad1 As Single, unPasXGrad2 As Single
    Dim unMaxXreel As Single, unMinXreel As Single
    
    EcrireDansFichier = True
    
    ' Active la routine de gestion d'erreur.
    On Error GoTo ErreurEcriture
    
    ' Fermeture du fichier pour délocké et ainsi pouvoir écrire dedans.
    If uneForm.monFichId <> 0 Then
        'Cas d'un Site qui n'est pas Sans Nom (Titre Etude + unNuméro)
        unFichId = uneForm.monFichId
        Close #unFichId
    End If
        
    'Ouvre le fichier en écriture.
    unFichId = FreeFile(0)
    uneForm.monFichId = unFichId
    Open unNomFich For Output As #unFichId
        
    'Remplissage du fichier à partir des données de l'itinéraire (=uneForm)
    '(cf Format de fichier MiTemps .mit)
    With uneForm
        'Ecriture de l'entête des fichiers *.mit
        Write #unFichId, "Fichier " + App.Title
        'Ecriture des libellés des conditions météo
        uneLigTexte = ""
        For i = 1 To uneForm.maColMeteo.Count - 1
            uneLigTexte = uneLigTexte + uneForm.maColMeteo(i) + ","
        Next i
        'Dernier libellé pas de virgule à la fin
        uneLigTexte = uneLigTexte + uneForm.maColMeteo(i)
        Write #unFichId, uneLigTexte
        'Ecriture des min et max total en distance (parcours complet sans section)
        Write #unFichId, .monMinDtot, .monMaxDtot
        
        'Calcul des max et min Temps réels pour les stocker
        unMaxXreel = .monMaxT
        unMinXreel = .monMinT
        'Calcul des pas de graduations primaires et secondaires et arrondis
        'de la valeur mini à la graduation secondaire juste inférieure
        'et de la valeur maxi à la graduation secondaire juste supérieure
        ArrondirXMinXMaxGrad2 unPasXGrad1, unPasXGrad2, unMaxXreel, unMinXreel
        .monMaxT = unMaxXreel
        .monMinT = unMinXreel
        
        'Stockage des max et min Vitesses réels pour les stocker
        unMaxXreel = .monMaxV
        unMinXreel = .monMinV
        'Calcul des pas de graduations primaires et secondaires et arrondis
        'de la valeur mini à la graduation secondaire juste inférieure
        'et de la valeur maxi à la graduation secondaire juste supérieure
        ArrondirXMinXMaxGrad2 unPasXGrad1, unPasXGrad2, unMaxXreel, unMinXreel
        .monMaxV = unMaxXreel
        .monMinV = unMinXreel
        
        'Ecriture des min et max  en distance, vitesse et temps
        Write #unFichId, .monMinD, .monMaxD, .monMinV, .monMaxV, .monMinT, .monMaxT
        'Ecriture du nom de l'itinéraire et de sa longueur
        Write #unFichId, .TextNomIti.Text, .TextLongIti.Text
        'Ecriture des données de la section de travail
        Write #unFichId, .CheckSection.Value, .ComboRepDebSec.ListIndex + 1, .ComboRepFinSec.ListIndex + 1
        'Ecriture du nombre de repères et de parcours (utile pour la lecture du fichier)
        'sur 4 caractères pour le nb de parcours car on écrase ces 4 caractères lors
        'd'un import de parcours dans les fichiers MIT ainsi on ne déborde jamais
        'sur les autres lignes en faisant un seek ici plus un print en mode ajout
        '(= Append, cf BtnImport_Click de la form frmImportMTB)
        uneLigTexte = Format(.maColRepere.Count) + "," + Format(.maColParcours.Count, "000#")
        Print #unFichId, uneLigTexte
        'Ecriture des données des repères
        For i = 1 To .maColRepere.Count
            Set unRep = .maColRepere(i)
            Write #unFichId, unRep.monNomLong, unRep.monNomCourt, unRep.monAbsCurv, unRep.monTypeIcone
        Next i
        'Ecriture des données des parcours, pour chaque parcours :
        'Ligne 1 : Les champs modifiables, Ligne 2 : les non-modifiables
        'stockés dans le spread parcours dont le nombre de repères topés
        'Ligne 3 : le nombre de pas de mesure,
        '          le pas de mesure en secondes,
        '          le premier et dernier pas de mesure
        'Ligne 4 à nb pas  modulo 13, (13 pas de mesure par ligne)
        'Aprés les N absCurv des repères topés (10 abscurv par ligne)
        'Puis les N Temps passage des repères topés (10 temps par ligne)
        For i = 1 To .maColParcours.Count
            Set unPar = .maColParcours(i)
            EcrireDonneesParcoursDansFichierMIT unFichId, unPar
        Next i
    End With
    
    'Mise à jour du titre de la fenetre étude courante
    uneForm.Caption = MsgIti0 + unNomFich
    
    'Fermeture du fichier.
    Close #unFichId
        
    'Ouverture du fichier en lock pour éviter deux ouvertures
    Open unNomFich For Input Lock Read Write As #unFichId
    
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    ' Quitte pour éviter le gestionnaire d'erreur.
    Exit Function
    
    ' Routine de gestion d'erreur qui évalue le numéro d'erreur.
ErreurEcriture:
    
    EcrireDansFichier = False
    ' Traite les autres situations ici...
    unMsg = MsgErreur + Format(Err.Number) + " : " + Err.Description
    MsgBox unMsg, vbCritical
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    'fermeture du fichier
    Close #unFichId
    'Ouverture du fichier en lock pour éviter deux ouvertures
    Open unNomFich For Input Lock Read Write As #unFichId
    Exit Function
End Function


Public Sub EcrireDonneesParcoursDansFichierMIT(unFichId As Byte, unPar As Parcours)
    'Procédure écrivant les données d'un parcours dans un fichier itinéraire
    'd'extension MIT
    'Ecriture des données des parcours, pour chaque parcours :
    'Ligne 1 : Les champs modifiables, Ligne 2 : les non-modifiables
    'stockés dans le spread parcours dont le nombre de repères topés
    'Ligne 3 : le nombre de pas de mesure,
    '          le pas de mesure en secondes,
    '          le premier et dernier pas de mesure
    'Ligne 4 à nb pas  modulo 13, (13 pas de mesure par ligne)
    'Aprés les N absCurv des repères topés (10 abscurv par ligne)
    'Puis les N Temps passage des repères topés (10 temps par ligne)
    Dim j As Long
    
    unNbRepTop = UBound(unPar.monTabAbsRep)
    Write #unFichId, unPar.monNom, unPar.monIsUtil, unPar.monIsParcoursMoyen, unPar.maCouleur, unPar.monEnqueteur, unPar.monNumVeh, unPar.maMeteo, unPar.maDate, unPar.monJourSemaine, unPar.monHeureDebut
    Write #unFichId, unNbRepTop, unPar.maDistPar, unPar.monTabAbsRep(unNbRepTop), unPar.maDuree, unPar.monCoefEta
    Write #unFichId, unPar.monNbPas, unPar.monPasMesure, unPar.monFirstPas, unPar.monLastPas
    
    'Ecriture des pas de mesure, 13 par ligne maxi, séparés par une virgule, sauf le dernier de chaque ligne
    'Chaque pas 5 car maxi, avec la virgule 6 car maxi ==> ligne de 78 = 6*13 car maxi, donc <= 80 lisible sur une page d'éditeur
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
        'ci-dessus k = unNbParLig aprés sortie du for
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
    
    'Ecriture des abs curv des repères topés, 10 par ligne maxi, séparés par une virgule, sauf le dernier de chaque ligne
    'Chaque abs 7 car maxi, avec la virgule 8 car maxi ==> ligne de 80=10*8 car maxi, donc <= 80 lisible sur une page d'éditeur
    unNbParLig = 10
    unQuot = unNbRepTop \ unNbParLig
    For j = 1 To unQuot
        uneLigTexte = ""
        For k = 1 To unNbParLig - 1
            uneLigTexte = uneLigTexte + Format(CLng(unPar.monTabAbsRep((j - 1) * unNbParLig + k))) + ","
        Next k
        uneLigTexte = uneLigTexte + Format(CLng(unPar.monTabAbsRep((j - 1) * unNbParLig + k)))
        'ci-dessus k = unNbParLig aprés sortie du for
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

    'Ecriture des temps de passage des repères topés, 10 par ligne maxi, séparés par une virgule, sauf le dernier de chaque ligne
    'Chaque temps 7 car maxi, avec la virgule 8 car maxi ==> ligne de 80=10*8 car maxi, donc <= 80 lisible sur une page d'éditeur
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
        'ci-dessus k = unNbParLig aprés sortie du for
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
    'Ouvre l'étude contenue dans le fichier passé en paramètre
    'pour récupérer les repères et parcours dans les deux collections
    'passées en paramètres
    'On retourne la ligne où se trouve les nombres de repères et de parcours
    'si retour = 0 la lecture de l'itinéraire ou étude a échoué
    
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
    
    'Initialisation de la valeur de retour à 0
    'si on va au bout elle vaudra la ligne contenant les nombres
    'de repères et de parcours, sinon 0
    RecupererContenuEtude = 0
    
    'Si protection invalide on ne fait rien
    'If ProtectCheck(2) <> 0 Then Exit Function
    
    'Lecture du fichier .mit
    ' Active la routine de gestion d'erreur.
    'MsgBox "Suppression du On Error GoTo ErreurLecture2"
    On Error GoTo ErreurLecture2
    
    'Ouverture du fichier en lecture lockée pour éviter deux ouvertures
    unFichId = FreeFile(0)
    Open unNomFich For Input Lock Read Write As #unFichId
       
    'Lecture de l'entête des fichiers *.mit
    Input #unFichId, uneString
    If uneString <> "Fichier " + App.Title Then
        'Cas d'un fichier qui n'est pas un fichier MiTemps
        '===> Fermeture du fichier.
        Close #unFichId
        MsgBox MsgErreur + MsgFileNotFile + App.Title + " version " + Format(App.Major) + "." + Format(App.Minor), vbCritical
        ' Désactive la récupération d'erreur.
        On Error GoTo 0
        Exit Function
    Else
        'Cas d'un fichier MiTemps *.Mit de la version 3.0
        '1ère ligne du fichier MIT = "Fichier MiTemps"
        
        'Récupération des libellés des conditions météo
        Input #unFichId, uneString
        For i = 0 To 7
            unePos = InStr(1, uneString, ",")
            uneString = Mid(uneString, unePos + 1)
        Next i
        
        'Récupération des min et max total en distance (parcours complet sans section)
        Input #unFichId, unTabSng(1), unTabSng(2)
        'Récupération des min et max en distance, vitesse et temps
        Input #unFichId, unMinD, unMaxD, unTabSng(3), unTabSng(4), unTabSng(5), unTabSng(6)
        
        'Récupération du nom de l'itinéraire et de sa longueur
        Input #unFichId, uneString, uneString1
        
        'Récupération des données de la section de travail
        Input #unFichId, unCheckSection, unIndRepDeb, unIndRepFin
        
        'Stockage de la position de lecture juste avant la lecture du nombre
        'de repères et parcours, on s'en sert lors de l'importation pour écrire
        'dans le fichier mit de l'itinéraire chargé le nombre de parcours + 1
        unNumLigNbPar = Seek(unFichId)
        
        'Récupération du nombre de repères et de parcours
        Input #unFichId, unNbRep, unNbPar
        'Récupération, stockage et création des repères
        For i = 1 To unNbRep
            Input #unFichId, uneString, uneString1, uneAbsCurv, unTypeIco
            Set unRep = uneColRep.Add(uneString, uneString1, uneAbsCurv, unTypeIco)
        Next i
        'Remise à zéro de unNbRep
        unNbRep = 0
        
        'Récupération des données des parcours, pour chaque parcours :
        'Ligne 1 : Les champs modifiables, Ligne 2 : les non-modifiables
        'stockés dans le spread parcours dont le nombre de repères topés
        'Ligne 3 : le nombre de pas de mesure,
        '          le pas de mesure en secondes,
        '          le premier et dernier pas de mesure
        'Ligne 4 à nb pas  modulo 13, (13 pas de mesure par ligne)
        'Aprés les N absCurv des repères topés (10 abscurv par ligne)
        'Puis les N Temps passage des repères topés (10 temps par ligne)
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
            
            'Stockage dans unNbRep du nombre de repères maxis
            If unNbRepTop > unNbRep Then unNbRep = unNbRepTop
            
            Input #unFichId, unNbPas, unLong1, unInt1, unInt2
            unPar.monNbPas = unNbPas
            unPar.monPasMesure = unLong1
            unPar.monFirstPas = unInt1
            unPar.monLastPas = unInt2
                        
            'Récupération des pas de mesure
            For j = 1 To unNbPas
                'Lecture de la donnée
                Input #unFichId, unTabSng(1)
                unPar.monTabDist(j) = unTabSng(1)
            Next j
                        
            'Allocation dynamique des tableaux liés aux repères topés
            unTabAbsRep = unPar.monTabAbsRep
            unTabTempsRep = unPar.monTabTempsRep
            ReDim unTabAbsRep(1 To unNbRepTop)
            ReDim unTabTempsRep(1 To unNbRepTop)
            
            'Récupération des abs curv des repères topés
            For j = 1 To unNbRepTop
                'Lecture de la donnée
                Input #unFichId, unLong1
                unTabAbsRep(j) = unLong1
            Next j
        
            'Récupération des temps de passage des repères topés
            For j = 1 To unNbRepTop
                'Lecture de la donnée
                Input #unFichId, unTabSng(1)
                unTabTempsRep(j) = unTabSng(1)
            Next j
            
            'Affectation des pointeurs sur les tableaux du parcours
            unPar.monTabAbsRep = unTabAbsRep
            unPar.monTabTempsRep = unTabTempsRep
        Next i
    End If
        
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
        
    'Vidage de tous les events en attente
    DoEvents
        
    ' Quitte pour éviter le gestionnaire d'erreur en retournant que tout
    's'est bien passé, donc le numéro de la ligne où on lit le nombre de parcours
    Close #unFichId
    RecupererContenuEtude = unNumLigNbPar
    Exit Function
    
    ' Routine de gestion d'erreur qui évalue le numéro d'erreur.
ErreurLecture2:
    
    ' Traite les autres situations ici...
    unMsg = MsgOpenError + unNomFich + Chr(13) + Chr(13) + MsgErreur + Format(Err.Number) + " : " + Err.Description
    If Err.Number = 70 Then unMsg = unMsg + " (" + UCase(MsgDejaOpen) + ")"
    MsgBox unMsg, vbCritical
    'Fermeture du fichier et retour 0 de la fonction car échec lors de la lecture
    RecupererContenuEtude = 0
    Close #unFichId
    ' Désactive la récupération d'erreur.
    On Error GoTo 0
    Exit Function
End Function

