Attribute VB_Name = "ModuleIHM"
Public Sub RetaillerOngletItiRef(uneForm As Form)
    'Retaillage de l'onglet Itinéraire de référence
    Dim i As Integer
    
    'Retaillage du spread des repères
    uneForm.SpreadRepere.Left = 60

    If uneForm.maColParcours.Count > 0 Then
        'Cas d'un fichier ayant des parcours affectés
        'Retaillage du spread des parcours
        uneForm.SpreadParcours.Left = uneForm.SpreadRepere.Left
        uneForm.SpreadParcours.Width = uneForm.TabData.Width - uneForm.SpreadParcours.Left * 2
        
        'Aggrandissement vertical des deux spreads repères et parcours
        'pour mieux occuper l'espace libre suivant la résolution et la
        'taille de la fenêtre fille
        unePlaceLibre = uneForm.TabData.Height - 120 - TSpreadPar - HSpreadPar
        uneForm.SpreadRepere.Height = HSpreadRep + unePlaceLibre / 2
        uneForm.SpreadParcours.Top = TSpreadPar + unePlaceLibre / 2
        uneForm.SpreadParcours.Height = HSpreadPar + unePlaceLibre / 2
        uneForm.BtnFiltreSel.Top = uneForm.SpreadRepere.Top + uneForm.SpreadRepere.Height + 120
        uneForm.BtnMeteo.Top = uneForm.BtnFiltreSel.Top
        uneForm.BtnSuppPar.Top = uneForm.BtnFiltreSel.Top
        uneForm.LabelInfoColor.Top = uneForm.BtnFiltreSel.Top + (uneForm.BtnFiltreSel.Height - uneForm.LabelInfoColor.Height) / 2
        
        'Récup de la Hauteur de la ligne 0 du spread parcours
        uneH0 = uneForm.SpreadParcours.RowHeight(0)
        uneH = uneForm.SpreadParcours.RowHeight(1)
        'Retaillage de spread parcours pour que l'ascenseur horizontal
        'soit proche de la dernière ligne, on rajoute 7.5% à chaque ligne
        'du spread en assimilant l'ascenseur à une ligne de plus, d'où le + 1
        uneNewH = uneH0 * 1.075 + uneH * 1.075 * (uneForm.SpreadParcours.MaxRows + 1)
        If uneNewH < uneForm.SpreadParcours.Height Then
            'Retaillage uniquement si la hauteur trouvé n'est pas plus grande
            'que allant jusqu'en bas de l'onglet
            uneForm.SpreadParcours.Height = uneNewH
        End If
        
        'Affichage du spread des parcours affectés
        uneForm.SpreadParcours.Visible = True
    Else
        'Cas d'un fichier n'ayant pas de parcours affectés
        'ou Cas d'une nouvelle étude
        uneForm.SpreadParcours.Visible = False
        uneForm.SpreadRepere.Height = uneForm.TabData.Height - uneForm.SpreadRepere.Top - 120
    End If
End Sub

Public Sub RetaillerOngletCbeDT(uneForm As Form, Optional unTestRedessin As Boolean = False)
    'Retaillage de l'onglet Courbe Distance/Temps
    If uneForm.GetTabRedOnglet(OngletCbeDT) = False Then
        'Pas de Redessin car pas de modif dans l'onglet ItiRef
        Exit Sub
    End If
    'Mis à faux pour ne redessiner qu'à la prochaine modif dans onglet ItiRef
    uneForm.SetTabRedOnglet OngletCbeDT, False
    
    With uneForm
        'Calage à gauche du spread d'info du parcours sélectionné
        .SpreadInfoParcoursDT.Top = .TabData.TabHeight + 90
        .SpreadInfoParcoursDT.Left = 75
        'Calcul pour mettre la taille maximun de la picture box permettant
        'de dessiner les courbes distance/temps
        .PicBoxDT.Left = .SpreadInfoParcoursDT.Left + .SpreadInfoParcoursDT.Width
        .PicBoxDT.Top = .TabData.TabHeight + PicBoxTop
        .PicBoxDT.Height = .TabData.Height - .TabData.TabHeight - PicBoxTop * 2
        .PicBoxDT.Width = .TabData.Width - .SpreadInfoParcoursDT.Width - PicBoxTop * 2
    
        'Redessin de la courbe distance/temps avec dessin en trait gros
        'du parcours sélectionné
        DessinerCourbe uneForm, .PicBoxDT, OngletCbeDT
    End With
End Sub

Public Sub RetaillerOngletCbeDV(uneForm As Form, Optional unTestRedessin As Boolean = False)
    'Retaillage de l'onglet Courbe Distance/Vitesse
    If uneForm.GetTabRedOnglet(OngletCbeDV) = False Then
        'Pas de Redessin car pas de modif dans l'onglet ItiRef
        Exit Sub
    End If
    'Mis à faux pour ne redessiner qu'à la prochaine modif dans onglet ItiRef
    uneForm.SetTabRedOnglet OngletCbeDV, False
    
    With uneForm
        'Calage à gauche du spread d'info du parcours sélectionné
        .SpreadInfoParcoursDV.Top = .TabData.TabHeight + 90
        .SpreadInfoParcoursDV.Left = 75
        'Calcul pour mettre la taille maximun de la picture box permettant
        'de dessiner les courbes distance/vitesses
        .PicBoxDV.Left = .SpreadInfoParcoursDV.Left + .SpreadInfoParcoursDV.Width
        .PicBoxDV.Top = .TabData.TabHeight + PicBoxTop
        .PicBoxDV.Height = .TabData.Height - .TabData.TabHeight - PicBoxTop * 2
        .PicBoxDV.Width = .TabData.Width - .SpreadInfoParcoursDV.Width - PicBoxTop * 2
    
        'Redessin de la courbe distance/tempsavec dessin en trait gros
        'du parcours sélectionné
        DessinerCourbe uneForm, .PicBoxDV, OngletCbeDV
    End With
End Sub

Public Sub RetaillerOngletSynoV(uneForm As Form, Optional unTestRedessin As Boolean = False)
    'Retaillage de l'onglet Synoptique des vitesses
    If uneForm.GetTabRedOnglet(OngletSynoV) = False Then
        'Pas de Redessin car pas de modif dans l'onglet ItiRef
        Exit Sub
    End If
    'Mis à faux pour ne redessiner qu'à la prochaine modif dans onglet ItiRef
    uneForm.SetTabRedOnglet OngletSynoV, False
    
    With uneForm
        'Calage à gauche de la frame des légendes de classes de vitesse
        .FrameLegende.Top = .TabData.TabHeight + 90
        .FrameLegende.Left = 75
        'Calcul pour mettre la taille maximun de la picture box permettant
        'de dessiner le synoptique des vitesses
        .PicBoxSynoV.Top = .TabData.TabHeight + PicBoxTop
        .PicBoxSynoV.Left = .FrameLegende.Left + .FrameLegende.Width
        .PicBoxSynoV.Height = .TabData.Height - .TabData.TabHeight - PicBoxTop * 2
        .PicBoxSynoV.Width = .TabData.Width - .FrameLegende.Width - PicBoxTop * 2
    
        'Redessin du synotique des vitesses
        DessinerSynoV uneForm, .PicBoxSynoV
    End With
End Sub

Public Sub RetaillerOngletHistV(uneForm As Form, Optional unTestRedessin As Boolean = False)
    'Retaillage de l'onglet Histogramme des vitesses
    If uneForm.GetTabRedOnglet(OngletHistV) = False Then
        'Pas de Redessin car pas de modif dans l'onglet ItiRef
        Exit Sub
    End If
    'Mis à faux pour ne redessiner qu'à la prochaine modif dans onglet ItiRef
    uneForm.SetTabRedOnglet OngletHistV, False
    
    With uneForm
        'Calage à gauche du MSCHART = histogramme des classes de vitesse
        .MSChart1.Top = .TabData.TabHeight + 90
        .MSChart1.Left = 75
        'Calcul pour mettre la taille maximun de la picture box permettant
        'de dessiner l'histogramme des vitesses
        .MSChart1.Height = .TabData.Height - .TabData.TabHeight - PicBoxTop * 2
        .MSChart1.Width = .TabData.Width - PicBoxTop * 2
    
        'Redessin du synotique des vitesses
        DessinerHistoV uneForm
        DoEvents
    End With
End Sub

Public Sub RetaillerOngletTabBr(uneForm As Form, Optional unTestRedessin As Boolean = False)
    'Retaillage de l'onglet Tableau brut
    If uneForm.GetTabRedOnglet(OngletTabBr) = False Then
        'Pas de Redessin car pas de modif dans l'onglet ItiRef
        Exit Sub
    End If
    'Mis à faux pour ne redessiner qu'à la prochaine modif dans onglet ItiRef
    uneForm.SetTabRedOnglet OngletTabBr, False
    
    uneForm.MousePointer = vbHourglass
    uneForm.BtnExportTabBrut.Top = uneForm.TabData.TabHeight + 90
    uneForm.BtnExportTabBrut.Left = uneForm.TabData.Width - uneForm.BtnExportTabBrut.Width - 90
    uneForm.SpreadTabBrut.Left = (uneForm.TabData.Width - uneForm.SpreadTabBrut.Width) / 2 '90
    uneForm.SpreadTabBrut.Top = uneForm.BtnExportTabBrut.Top + uneForm.BtnExportTabBrut.Height + 90 'uneForm.TabData.TabHeight + 90
    uneForm.SpreadTabBrut.Height = uneForm.TabData.Height - uneForm.SpreadTabBrut.Top - 90 '- uneForm.TabData.TabHeight - 180 ' 180 = 90 * 2
    RemplirTabBrut uneForm
    uneForm.MousePointer = vbDefault
End Sub

Public Sub RetaillerOngletTabSS(uneForm As Form, Optional unTestRedessin As Boolean = False)
    'Retaillage de l'onglet Tableau Synthèse et statistiques
    If uneForm.GetTabRedOnglet(OngletTabSS) = False Then
        'Pas de Redessin car pas de modif dans l'onglet ItiRef
        Exit Sub
    End If
    'Mis à faux pour ne redessiner qu'à la prochaine modif dans onglet ItiRef
    uneForm.SetTabRedOnglet OngletTabSS, False
    
    uneForm.MousePointer = vbHourglass
    uneForm.BtnExportTabSS.Top = uneForm.TabData.TabHeight + 90
    uneForm.BtnExportTabSS.Left = uneForm.TabData.Width - uneForm.BtnExportTabSS.Width - 90
    uneForm.SpreadTabSS.Left = (uneForm.TabData.Width - uneForm.SpreadTabSS.Width) / 2
    uneForm.SpreadTabSS.Top = uneForm.BtnExportTabSS.Top + uneForm.BtnExportTabSS.Height + 90
    uneForm.SpreadTabSS.Height = uneForm.TabData.Height - uneForm.SpreadTabSS.Top - 90
    RemplirTabSS uneForm
    uneForm.MousePointer = vbDefault
End Sub


Public Function DonnerDistEcran(uneDistReel As Single, uneDistMaxReel As Single, uneDistMaxEcran As Single) As Single
    'Retourne la conversion d'une distance réelle en distance écran
    DonnerDistEcran = uneDistReel / uneDistMaxReel * uneDistMaxEcran
End Function

Public Function ConvertirEnEcran(unOrigEcran As Single, uneDistReel As Single, uneDistMaxReel As Single, uneDistMaxEcran As Single) As Single
    'Retourne la conversion d'une coordonnée réelle en coordonnées X ou Y à l'écran ou imprimante
    ConvertirEnEcran = unOrigEcran + uneDistReel / uneDistMaxReel * uneDistMaxEcran
End Function

Public Function ConvertirEnReel(unOrigReel As Single, uneDistEcran As Single, uneDistMaxReel As Single, uneDistMaxEcran As Single) As Single
    'Retourne la conversion d'une coordonnée écran réelle en coordonnées X ou Y réelle
    ConvertirEnReel = unOrigReel + uneDistEcran / uneDistMaxEcran * uneDistMaxReel
End Function

Public Function DonnerDistReel(uneDistEcran As Single, uneDistMaxReel As Single, uneDistMaxEcran As Single) As Single
    'Retourne la conversion d'une distance écran en distance réelle
    DonnerDistReel = uneDistEcran / uneDistMaxEcran * uneDistMaxReel
End Function


Public Sub RemplirSpreadInfoParcoursSel(unSpreadInfo As vaSpread, uneForm As Form, unIndParcours As Integer)
    'Remplir le spread d'info du bon onglet du parcours sélectionné de la form
    'passé en paramètre (= fenêtre itinéraire)
    Dim unParcours As Parcours
    
    If unIndParcours = 0 Then
        'Pas de parcours sélectionné
        unRemplirVide = True
    Else
        Set unParcours = uneForm.maColParcours(unIndParcours)
        If unParcours.monIsUtil Then
            'Cas où le parcours sélectionné est utilisé donc visible
            unRemplirVide = False
        Else
            unRemplirVide = True
        End If
    End If
    
    'Positionnement en col 1 car c'est la seule colonne des spread info
    unSpreadInfo.Col = 1
    If unRemplirVide Then
        'Remplissage à vide
        For i = 1 To unSpreadInfo.MaxRows
            unSpreadInfo.Row = i
            unSpreadInfo.Text = ""
            unSpreadInfo.BackColor = vbInfoBackground
        Next i
    Else
        'Remplissage avec les info du parcours sélectionné
        'Affichage du nom
        unSpreadInfo.Row = 1
        unSpreadInfo.Text = unParcours.monNom
        'Affichage de la couleur du parcours en couleur de fond
        unSpreadInfo.Row = 2
        unSpreadInfo.BackColor = unParcours.maCouleur
        'Affichage de la date de mesure
        unSpreadInfo.Row = 3
        unSpreadInfo.Text = unParcours.monJourSemaine + " " + Format(unParcours.maDate)
        'Affichage de l'heure de début de mesure
        unSpreadInfo.Row = 4
        unSpreadInfo.Text = unParcours.monHeureDebut
        'Affichage de la vitesse moyenne
        unSpreadInfo.Row = 5
        unSpreadInfo.Text = unParcours.maVmoy
        'Affichage de la vitesse mini
        unSpreadInfo.Row = 6
        unSpreadInfo.Text = unParcours.maVmin
        'Affichage de la vitesse maxi
        unSpreadInfo.Row = 7
        unSpreadInfo.Text = unParcours.maVmax
        'Affichage du nombre d'arrêts
        unSpreadInfo.Row = 8
        unSpreadInfo.Text = unParcours.monNbArret
        'Affichage de la durée totale d'arrêts
        unSpreadInfo.Row = 9
        unSpreadInfo.Text = FormatterTempsEnHMNS(unParcours.monTpsArret)
        'Affichage du nombre de double tops
        unSpreadInfo.Row = 10
        unSpreadInfo.Text = unParcours.monNbDbTop
        'Affichage de la durée totale des double tops
        unSpreadInfo.Row = 11
        unSpreadInfo.Text = FormatterTempsEnHMNS(unParcours.monTpsDbTop)
        'Affichage de la durée du parcours total ou sur la section de travail
        unSpreadInfo.Row = 12
        unSpreadInfo.Text = FormatterTempsEnHMNS(unParcours.monTFinSection - unParcours.monTDebSection)
        'Affichage de la distance parcourue sur le parcours total
        'ou sur la section de travail en mètres, on ne multiplie pas par le
        'coef d'étalonnage car la distance a été obtenu par cumul de distances
        'multipliée par ce coef d'étalonnage
        unSpreadInfo.Row = 13
        unSpreadInfo.Text = Format(CLng(unParcours.maDistParSection / 10), "#,###,###")
    End If
End Sub

Public Function CreerRepere(uneFrmD As frmDocument, unNomLong As String, unNomCourt As String, uneAbsCurv As Long, unTypeIcone As Byte) As Repere
    'Création de l'instance de la classe Repere
    'création de son icône de visualisation et
    'de sa ligne dans le tableau des repères de l'onglet Iti Référence
    'Mis en dernière ligne dans le tableau des repères
    'Ajout dans les combobox de début et fin de section
    Dim unRepere As Repere, uneCleCol As String
    Dim uneIcone As ListImage
    
    'Pour avoir un clé unique dans la collection des repères,
    'pour ce repère on incrémente le nombre de repères
    uneFrmD.monNbRepere = uneFrmD.monNbRepere + 1
    uneCleCol = "Rep" + Format(uneFrmD.monNbRepere)
    
    'Vérification de l'unicité du nom court
    While VerifierNomCourtUnique(uneFrmD, unNomCourt) = False
        unNomCourt = "Repere " + Format(uneFrmD.monNbRepere + 1)
        'unNomLong = unNomCourt
    Wend
    
    'Création de l'instance de la classe Repere
    Set unRepere = uneFrmD.maColRepere.Add(unNomLong, unNomCourt, uneAbsCurv, unTypeIcone, uneCleCol)
    
    'Création de son icône de visualisation,
    'toujours le dernier créé d'où le .count
    'et des liens entre icone et objet repère
    Load uneFrmD.ImageRepere(uneFrmD.monNbRepere)
    'Pointeur de l'objet Repere vers son icone
    Set unRepere.monIcone = uneFrmD.ImageRepere(uneFrmD.monNbRepere)
    'Pointeur de l'icone vers son objet Repere par la clé de collection
    'stocké dans son tag
    uneFrmD.ImageRepere(uneFrmD.monNbRepere).Tag = uneCleCol
    
    'Création et remplissage de la ligne dans le tableau des
    'repères de l'onglet Itinéraire de Référence
    '==> Ajout en dernière ligne du spread SpreadRepere
    uneFrmD.SpreadRepere.MaxRows = uneFrmD.SpreadRepere.MaxRows + 1
    uneFrmD.SpreadRepere.Row = uneFrmD.SpreadRepere.MaxRows
    uneFrmD.SpreadRepere.Col = 1
    uneFrmD.SpreadRepere.Text = unNomLong
    uneFrmD.SpreadRepere.Col = 2
    uneFrmD.SpreadRepere.Text = unNomCourt
    uneFrmD.SpreadRepere.Col = 3
    uneFrmD.SpreadRepere.Text = Format(uneAbsCurv)
    uneFrmD.SpreadRepere.Col = 4
    'Récupération de l'image représentant le repère
    Set uneIcone = DonnerIconeRepere(unTypeIcone)
    uneFrmD.SpreadRepere.TypePictPicture = uneIcone.Picture
    uneFrmD.SpreadRepere.Col = 5
    uneFrmD.SpreadRepere.TypeComboBoxCurSel = unTypeIcone - 1
    
    'MAJ des liens entre l'objet repère et sa ligne
    'dans le spread des repères (dans la cellule de la dernière colonne)
    uneFrmD.SpreadRepere.Col = uneFrmD.SpreadRepere.MaxCols
    uneFrmD.SpreadRepere.Text = uneCleCol
    
    'Mettre à jour l'info bulle et l'icône de l'icône repère
    uneFrmD.ImageRepere(uneFrmD.monNbRepere).ToolTipText = unNomCourt + " / Type : " + uneIcone.Tag + " / AbsCurv = " + Format(unRepere.monAbsCurv) + " m"
    uneFrmD.ImageRepere(uneFrmD.monNbRepere).Picture = uneIcone.Picture
    
    'Remplissage des combobox listant les début et fin de section possibles
    uneFrmD.ComboRepDebSec.AddItem unRepere.monNomCourt
    uneFrmD.ComboRepFinSec.AddItem unRepere.monNomCourt

    'Rendre active la ligne créée du spread, c'est tjs la dernière
    uneFrmD.SpreadRepere.Row = uneFrmD.SpreadRepere.MaxRows
    uneFrmD.SpreadRepere.Col = 1
    uneFrmD.SpreadRepere.Action = 0 'SS_ACTION_ACTIVE_CELL
    
    'On retourne le repère créé
    Set CreerRepere = unRepere
End Function

Public Function DonnerIconeRepere(unTypeIcone As Byte) As ListImage
    'Retourne l'image représentant le repère
    'ayant un icone du type unTypeIcone
    'stockée dans les images de l'imaglist ListIcons
    'de la fenêtre MDI mère dont la variable globale est frmMain
    Dim unNbIconRep As Integer
    
    unNbIconRep = frmMain.ListIcons.ListImages.Count
    If unTypeIcone > 0 And unTypeIcone < unNbIconRep + 1 Then '17 Then
        Set DonnerIconeRepere = frmMain.ListIcons.ListImages(unTypeIcone)
    Else
        'Autres valeurs non comprises entre 1 et 16
        'qui sont les valeurs des images possibles de l'imaglist ListIcons
        'de la fenêtre MDI mère ==> Erreur de programmation.
        MsgBox MsgErreurProg + MsgErreurTypeIconeInconnu + MsgIn + "ModuleIHM:DonnerIconeRepere", vbCritical
        'Par défaut on met l'icone divers (le 15 ème position, triangle jaune)
        Set DonnerIconeRepere = frmMain.ListIcons.ListImages(15)
    End If
End Function

Public Sub DessinerRepere(uneFrmD As frmDocument, unRep As Repere, Optional unAncienAbsCurv As Long = -1000)
    'Dessin de l'icône du repere dans la frame de droite verticale
    'entre le min et le max en ditance des autres repères de l'itinéraire
    'Si l'abscisse curviligne (= distance) n'est pas entre le min
    'et le max, on modifie l'un ou l'autre et on redessine avec le
    'nouveau zoom déduit de ce nouveau min ou max
    Dim unMaxYecran As Single, uneHtPicBox As Single
    
    'Récupération de la position écran y max
    'pour être au même niveau dans la frame vertical des repères
    'et les distances de courbes DV et DT
    uneHtPicBox = uneFrmD.TabData.Height - uneFrmD.TabData.TabHeight - PicBoxTop * 2
    uneFrmD.maDistMaxEcranY = uneHtPicBox - PicBoxMargeH - PicBoxMargeB
    unMaxYecran = uneFrmD.TabData.TabHeight + PicBoxTop + PicBoxMargeH
    
    unRep.monIcone.Visible = True
    
    If unRep.monAbsCurv >= uneFrmD.monMinD And unRep.monAbsCurv <= uneFrmD.monMaxD Then
        'Cas où le nouveau repère est entre le min et le max
        'On place à la bonne échelle écran
        unYecran = ConvertirEnEcran(unMaxYecran, uneFrmD.monMaxD - unRep.monAbsCurv, uneFrmD.monMaxD - uneFrmD.monMinD, uneFrmD.maDistMaxEcranY)
        unRep.monIcone.Top = unYecran - unRep.monIcone.Height / 2
        uneFrmD.PictureItiRef.Line (0, unYecran)-(unRep.monIcone.Left, unYecran)
        'Effacement de l'ancienne ligne de rappel correspondant
        'à l'abscisse curviligne avant modif en la redessinant de la
        'couleur du background
        unYecran = ConvertirEnEcran(unMaxYecran, uneFrmD.monMaxD - unAncienAbsCurv, uneFrmD.monMaxD - uneFrmD.monMinD, uneFrmD.maDistMaxEcranY)
        uneFrmD.PictureItiRef.Line (0, unYecran)-(unRep.monIcone.Left, unYecran), uneFrmD.PictureItiRef.BackColor
    Else
        'Cas où le nouveau repère n'est pas entre le min et le max
        'On modifie le nouveau de zoom
        'On place à la nouvelle bonne échelle écran
        'ainsi que tous les autres repères
        
        If uneFrmD.CheckSection.Value = 0 Then
            'Si on n'est pas en section de travail
            If unRep.monAbsCurv < uneFrmD.monMinD Then
                'Modif du min distance de section
                uneFrmD.monMinD = unRep.monAbsCurv
                If unRep.monAbsCurv < uneFrmD.monMinDtot Then
                    'Modif du min distance total
                    uneFrmD.monMinDtot = unRep.monAbsCurv
                End If
            Else
                'Modif du max distance de section
                uneFrmD.monMaxD = unRep.monAbsCurv
                If unRep.monAbsCurv > uneFrmD.monMaxDtot Then
                    'Modif du max distance total
                    uneFrmD.monMaxDtot = unRep.monAbsCurv
                End If
            End If
        End If
        
        'Modif de l'affichage de la longueur
        'uneFrmD.TextLongIti.Text = Format(uneFrmD.monMaxDtot - uneFrmD.monMinDtot)
        uneFrmD.TextLongIti.Text = Format(DonnerLongIti(uneFrmD))
        
        'Redessin total au bon zoom englobant entre minD et maxD
        RedessinerZoomTout uneFrmD
    End If
End Sub

Public Function FixerMargesImprimante(uneMargeG As Single, uneMargeD As Single, uneMargeH As Single, uneMargeB As Single) As Single
    'Retourne la largeur maximale des noms courts repères imprimés
    uneMargeG = 1.5 * UnCmEnTwips ' = 1.5cm
    'uneMargeD = 2.5 * UnCmEnTwips + Printer.TextWidth("WWWWWWWWWW") ' = 2.5cm + largeur de 10 W (= nom court maximun de repère)
    
    'Recherche du plus grand nom court pour bien cadrer la marge droite
    'dans l'itinéraire courant
    unMaxWidth = Printer.TextWidth("W") 'Initialisation du maximun
    For i = 1 To monIti.maColRepere.Count
        If Printer.TextWidth(monIti.maColRepere(i).monNomCourt) > unMaxWidth Then
           unMaxWidth = Printer.TextWidth(monIti.maColRepere(i).monNomCourt)
        End If
    Next i
    
    'Marge droite = 2.5cm + Largeur nom court maximun de repère + largeur icone repere
    uneMargeD = 2.5 * UnCmEnTwips + unMaxWidth + monIti.ImageRepere(0).Width
    
    uneMargeH = 1.5 * UnCmEnTwips ' = 1.5cm
    If uneMargeH < monIti.maMargeHaut Then uneMargeH = monIti.maMargeHaut
    uneMargeB = 2.5 * UnCmEnTwips   ' = 2.5cm
    
    FixerMargesImprimante = unMaxWidth
End Function

Public Sub FixerMargesPicBox(uneForm As Form, uneZoneDessin As Object, uneMargeG As Single, uneMargeD As Single, uneMargeH As Single, uneMargeB As Single)
    uneMargeG = 600 '540 = Largeur en twips du label autosizé valant 999999 (distance maxi en m)
    uneMargeD = 600 '360 = Largeur en twips du label autosizé valant 9999 (24h = 1440 mn)
    uneMargeH = PicBoxMargeH '195 = Hauteur en twips du label autosizé valant 999999
    uneMargeB = PicBoxMargeB '195 = Hauteur en twips du label autosizé valant 9999
    uneForm.maDistMaxEcranX = uneZoneDessin.Width - uneMargeG - uneMargeD
    uneForm.maDistMaxEcranY = uneZoneDessin.Height - uneMargeH - uneMargeB
End Sub


Public Sub AfficherNouveauRepere(uneForm As Form)
    'Création et affichage d'un nouveau repère
    'dans la fenêtre fille courante
    Dim unNbRep As Integer, uneLongIti As Long
    Dim unRep As Repere, uneRow As Integer
    
    'Indication d'une modif
    uneForm.maModif = True
    
    unNbRep = uneForm.monNbRepere
    'monNbRepere est incrémenté dans CreerRepere,
    'appeler plus bas dans cette procédure
    If unNbRep = 0 Then
        'Cas du premier repére créé, on le met en 0 par défaut
        uneLongIti = 0
    Else
        'Cas des autres repères, on les met par défaut 500 mètres
        'plus loin que le maxi en distance de l'itinéraire
        uneLongIti = uneForm.monMaxD + 500
        uneForm.TextLongIti.Text = Format(uneLongIti)
    End If
    
    'Création du repère, retourne nothing si création impossible
    uneRow = uneForm.SpreadRepere.ActiveRow 'Stockage de la ligne active car modifié dans CreerRepere
    Set unRep = CreerRepere(uneForm, "Repere " + Format(unNbRep + 1), "Repere " + Format(unNbRep + 1), uneLongIti, Autre)
    If unRep Is Nothing Then Exit Sub
    
    'Dessin du repère
    DessinerRepere uneForm, unRep
    
    'Sélection de la ligne dans le spread des repère
    If uneRow > 0 Then DeselectionnerRepere uneForm, uneRow
    SelectionnerRepere uneForm, uneForm.SpreadRepere.ActiveRow
End Sub

Public Sub SupprimerRepere(uneForm As Form, unNumRow As Integer)
    'Suppression du repère d'index unNumRow dans la collection
    'maColRepere de la form uneForm
    'En effet, l'index dans la collection vaut le numéro de ligne
    'dans le spread repere
    Dim unRep As Repere, unMsg As String
    Dim uneAbsCurv As Long, uneAbsCurvInf As Long, uneAbsCurvSup As Long
    Dim unYRepMin As Long, unYRepMax As Long
    
    'Suppression interdite si moins trois repères
    If uneForm.maColRepere.Count < 3 Then
        unMsg = "La suppression du repère début ou fin est interdite, si ce sont les seuls repères existants."
        unMsg = unMsg + Chr(13) + Chr(13) + "Modifier plutôt leurs propriétés dans l'onglet " + uneForm.TabData.TabCaption(OngletItiRef) + "."
        MsgBox unMsg, vbInformation
        Exit Sub
    End If
    
    'Récupération de la clé unique d'identification du repère
    'stockée dans la dernière colonne et positionnement sur cette ligne
    uneForm.SpreadRepere.Col = uneForm.SpreadRepere.MaxCols
    uneForm.SpreadRepere.Row = unNumRow
    uneCle = uneForm.SpreadRepere.Text
    
    'Récup du repère par sa clé unique d'identification du repère
    'dans la collection des repères de la form
    Set unRep = uneForm.maColRepere(uneCle)
    
    'Suppression interdite si le repère est celui de début
    'ou de fin de section éventuelle
    If unNumRow = uneForm.ComboRepDebSec.ListIndex + 1 Or unNumRow = uneForm.ComboRepFinSec.ListIndex + 1 Then
        If unNumRow = uneForm.ComboRepDebSec.ListIndex + 1 Then
            unePosition = "début"
        Else
            unePosition = "fin"
        End If
        unMsg = "La suppression du repère début ou fin de la section de travail éventuelle est interdite."
        unMsg = unMsg + Chr(13) + Chr(13) + "Après avoir cocher " + Chr(34) + uneForm.CheckSection.Caption + Chr(34) + ", modifier le repère " + unePosition + " de section pour pouvoir supprimer le repère " + Chr(34) + unRep.monNomCourt + Chr(34) + "."
        MsgBox unMsg, vbInformation
        Exit Sub
    End If
    
    'Demande de confirmation de suppression
    unMsg = "Voulez-vous vraiment supprimer le repère " + Chr(34) + unRep.monNomCourt + Chr(34) + " ?"
    If MsgBox(unMsg, vbQuestion + vbYesNo, "Confirmation de suppression") = vbNo Then Exit Sub
    
    'Indication d'une modif
    uneForm.maModif = True
    
    'Suppression de l'icone du repère et
    'effacement de sa ligne de rappel
    unYecran = unRep.monIcone.Top + unRep.monIcone.Height / 2
    uneForm.PictureItiRef.Line (0, unYecran)-(unRep.monIcone.Left, unYecran), uneForm.PictureItiRef.BackColor
    Unload unRep.monIcone
    
    'Suppression de la ligne dans le spread des repères
    uneForm.SpreadRepere.Action = 5 ' = SS_ACTION_DELETE_ROW
    'Suppression de la dernière ligne car une ligne vide est rajouté
    'lors d'une action delete row
    uneForm.SpreadRepere.MaxRows = uneForm.SpreadRepere.MaxRows - 1
    
    'Supression dans les combobox de début et fin de section
    'Suppression dans les deux listes égales des
    'combobox ComboRepDebSec et ComboRepFinSec
    unePos = unNumRow - 1
    uneForm.ComboRepDebSec.RemoveItem unePos
    uneForm.ComboRepFinSec.RemoveItem unePos
    
    'Stockage de l'abs curviligne pour savoir si la suppression
    'change le zoom
    uneAbsCurv = unRep.monAbsCurv
    uneAbsCurvInf = DonnerValGrad(uneForm, uneAbsCurv, 0) 'Arrondi à la graduation juste inférieure
    uneAbsCurvSup = DonnerValGrad(uneForm, uneAbsCurv, 1) 'Arrondi à la graduation juste supérieure
    
    'Suppression dans la collection des repères
    uneForm.maColRepere.Remove uneCle
    
    'Redessin éventuel total au bon zoom entre minD et maxD
    'La suppression n'est possible qu'en section de travail non définie
    '==> Min = MinTot et Max = MaxTot
    If uneAbsCurvInf <= uneForm.monMinD Then
        unYRepMin = DonnerYRepMin(uneForm)
        uneForm.monMinD = unYRepMin
        uneForm.monMinDtot = unYRepMin
        RedessinerZoomTout uneForm
    ElseIf uneAbsCurvSup >= uneForm.monMaxD Then
        unYRepMax = DonnerYRepMax(uneForm)
        uneForm.monMaxD = unYRepMax
        uneForm.monMaxDtot = unYRepMax
        RedessinerZoomTout uneForm
    End If
    
    'Modif de l'affichage de la longueur
    'uneForm.TextLongIti.Text = Format(uneForm.monMaxDtot - uneForm.monMinDtot)
    uneForm.TextLongIti.Text = Format(DonnerLongIti(uneForm))
    
    'Sélectionner le nouveau dernier repère
    SelectionnerRepere uneForm, uneForm.SpreadRepere.MaxRows
End Sub

Public Sub SelectionnerRepere(uneForm As Form, uneRow As Integer, Optional uneCol As Integer = 1)
    'Sélection du repère de la ligne uneRow du spread repere
    'grâce à la Clé d'identification du repère contenue
    'dans la dernière de la ligne active
    Dim unRep As Repere, uneImage As Image
    Dim unObj As Object, unMsg As String
    
    uneForm.SpreadRepere.Row = uneRow
    uneForm.SpreadRepere.Col = uneForm.SpreadRepere.MaxCols
    uneCle = uneForm.SpreadRepere.Text
    Set unRep = uneForm.maColRepere(uneCle)
    'Sélection graphique ==> apparition d'une bordure en premier plan
    unRep.monIcone.BorderStyle = vbFixedSingle
    'Mise au premier plan du repère sélectionné
    Set uneImage = unRep.monIcone
    uneImage.ZOrder 0
    DoEvents
    'On rend actif la ligne uneRow du spread repere
    uneForm.SpreadRepere.Row = uneRow
    uneForm.SpreadRepere.Col = uneCol
    uneForm.SpreadRepere.Action = 0 'SS_ACTION_ACTIVE_CELL
    'Stockage dans le tag de la fenêtre fille de la clé d'identification
    'correspond à ce repère
    uneForm.Tag = uneCle
    'Message d'info si on sélectionne dans le tableau des repères
    'un repère qui n'est pas dans la section de travail
    If uneForm.CheckSection.Value = 1 Then
        If unRep.monAbsCurv < uneForm.monMinD Or unRep.monAbsCurv > uneForm.monMaxD Then
            unMsg = "Le repère " + unRep.monNomCourt + " n'est pas dans la section de travail entre les repères "
            unMsg = unMsg + uneForm.ComboRepDebSec.Text + " et " + uneForm.ComboRepFinSec.Text + ", donc il n'est pas visible sur l'axe des distances situé à droite."
            MsgBox unMsg, vbInformation
        End If
    End If
    uneForm.SpreadRepere.SetFocus 'pour mettre le focus dans le spread repères
    uneForm.SpreadRepere.Refresh
End Sub

Public Sub DeselectionnerRepere(uneForm As Form, uneRow As Integer)
    'Désélection du repère de la ligne uneRow du spread repere
    'grâce à la Clé d'identification du repère contenue
    'dans la dernière de la ligne active
    uneForm.SpreadRepere.Row = uneRow
    uneForm.SpreadRepere.Col = uneForm.SpreadRepere.MaxCols
    uneCle = uneForm.SpreadRepere.Text
    If uneCle = "" Then Exit Sub
    Set unRep = uneForm.maColRepere(uneCle)
    'Désélection graphique ==> disparition de la bordure
    unRep.monIcone.BorderStyle = vbBSNone
End Sub


Public Sub RedessinerZoomTout(uneFrmD As frmDocument)
    'Redessin total au bon zoom englobant entre minD et maxD de la form
    Dim unMaxYecran As Single
    Dim uneDistMaxReelY As Single
    Dim unPasYGrad1 As Long, unPasYGrad2 As Long
    Dim uneMargeG As Single, uneMargeD As Single
    Dim uneMargeH As Single, uneMargeB As Single
    Dim unMaxYreel As Single, unMinYreel As Single
    
    'Initialisation des indicateurs de redessin des onglets de 1 à 6
    'à vrai pour déclencher le dessin lors de leur activation
    IndiquerToutRedessiner uneFrmD
    
    'Récupération de la position écran y max
    'pour être au même niveau dans la frame vertical des repères
    'et les distances de courbes DV et DT
    uneHtPicBox = uneFrmD.TabData.Height - uneFrmD.TabData.TabHeight - PicBoxTop * 2
    uneFrmD.maDistMaxEcranY = uneHtPicBox - PicBoxMargeH - PicBoxMargeB
    'Taille maxi de l'écran en Y
    unMaxYecran = uneFrmD.TabData.TabHeight + PicBoxTop + PicBoxMargeH
    
    'Calcul des pas de graduations primaires et secondaires et arrondis
    'de la valeur mini à la graduation secondaire juste inférieure
    'et de la valeur maxi à la graduation secondaire juste supérieure
    unMaxYreel = uneFrmD.monMaxD
    unMinYreel = uneFrmD.monMinD
    TrouverPasGradEtModifierMinMax unPasYGrad1, unPasYGrad2, unMaxYreel, unMinYreel
    uneFrmD.monMaxD = unMaxYreel
    uneFrmD.monMinD = unMinYreel
    'Stockage des pas de graduations en distance
    uneFrmD.monPasGrad1 = unPasYGrad1
    uneFrmD.monPasGrad2 = unPasYGrad2
    
    'Conversion en y écran des abscisses curvilignes de tous les
    'repères de la form grâce au nouveau niveau zoom englobant
    'entre le Min et le Max
    uneDistMaxReelY = uneFrmD.monMaxD - uneFrmD.monMinD
    uneFrmD.PictureItiRef.Cls
    For i = 1 To uneFrmD.maColRepere.Count
        If uneFrmD.maColRepere(i).monAbsCurv >= uneFrmD.monMinD And uneFrmD.maColRepere(i).monAbsCurv <= uneFrmD.monMaxD Then
            unYecran = ConvertirEnEcran(unMaxYecran, uneFrmD.monMaxD - uneFrmD.maColRepere(i).monAbsCurv, uneDistMaxReelY, uneFrmD.maDistMaxEcranY)
            uneFrmD.maColRepere(i).monIcone.Top = unYecran - uneFrmD.maColRepere(i).monIcone.Height / 2
            uneFrmD.PictureItiRef.Line (0, unYecran)-(uneFrmD.maColRepere(i).monIcone.Left, unYecran)
            uneFrmD.maColRepere(i).monIcone.Visible = True
        Else
            uneFrmD.maColRepere(i).monIcone.Visible = False
        End If
    Next i
End Sub

Public Sub ZoomToutSection(uneFrmD As frmDocument, unY1 As Long, unY2 As Long, unIndexRep As Integer)
    'Redessin des repères en niveau de zoom correspondant à l'englobant
    'contenant toute la section de travail entre début et fin
    'et sélection du repère de la ligne unIndexRep
    If unY1 < unY2 Then
        uneFrmD.monMaxD = unY2
        uneFrmD.monMinD = unY1
    Else
        uneFrmD.monMaxD = unY1
        uneFrmD.monMinD = unY2
    End If
    RedessinerZoomTout uneFrmD
    
    'Sélection du repère de début de section
    DeselectionnerRepere uneFrmD, uneFrmD.SpreadRepere.ActiveRow
    SelectionnerRepere uneFrmD, unIndexRep
End Sub

Public Sub ModifierIconeRepere(uneFrmD As frmDocument, unRep As Repere, Optional unTypeIco As Byte = 0)
    'Modif de l'icône du repère dans les colonnes 4 et 5 du spread repere
    'de la fenêtre fille (= itinéraire) et sur l'axe des distances
    Dim uneIcone As ListImage, unTypeIcone As Byte
    
    'Positionnement dans la cellule active du spread repère
    uneFrmD.SpreadRepere.Row = uneFrmD.SpreadRepere.ActiveRow
    uneFrmD.SpreadRepere.Col = uneFrmD.SpreadRepere.ActiveCol
    'Récupération de l'image représentant le repère
    If unTypeIco = 0 Then
        'Cas du remplacement  par click dans lespread des repères
        unTypeIcone = uneFrmD.SpreadRepere.TypeComboBoxCurSel + 1
    Else
        'Cas du changement en icone double top lors de leur détection
        'à la lecture d'un MTB
        unTypeIcone = unTypeIco
    End If
    Set uneIcone = DonnerIconeRepere(unTypeIcone)
    'Modif de l'image en colonne 4
    uneFrmD.SpreadRepere.Col = 4
    uneFrmD.SpreadRepere.TypePictPicture = uneIcone.Picture
    'Modif du libellé dans la combobox en colonne 5
    'listant les types d'icones
    uneFrmD.SpreadRepere.Col = 5
    uneFrmD.SpreadRepere.TypeComboBoxCurSel = unTypeIcone - 1
    'Modif de l'icone du repère sur l'axe des distance
    unRep.monIcone.Picture = uneIcone.Picture
    'Modif de l'entier donnant le type d'icône
    unRep.monTypeIcone = unTypeIcone
    'Modif de l'info-bulle
    unRep.monIcone.ToolTipText = unRep.monNomCourt + " / Type : " + uneIcone.Tag + " / AbsCurv = " + Format(unRep.monAbsCurv) + " m"
End Sub

Public Sub ModifierAbsCurvRepere(uneFrmD As frmDocument, unRep As Repere, unNewAbsCurv As Long)
    'Modif de l'abscisse curviligne du repère dans sa colonne du spread
    'repere de la fenêtre fille (= itinéraire) et sur l'axe des distances
    Dim unAncienAbsCurv As Long, unOldMinD As Long, unOldMaxD As Long
    Dim unAncienAbsCurvSup As Long, unAncienAbsCurvInf As Long
    Dim unYRepMin As Long, unYRepMax As Long
    
    If unRep.monAbsCurv <> unNewAbsCurv Then
        'Cas où l'abscisse curviligne change
        unAncienAbsCurv = unRep.monAbsCurv
        unAncienAbsCurvSup = DonnerValGrad(uneFrmD, unRep.monAbsCurv, 1) 'Arrondi à la graduation supérieure
        unAncienAbsCurvInf = DonnerValGrad(uneFrmD, unRep.monAbsCurv, 0) 'Arrondi à la graduation inférieure
        unRep.monAbsCurv = unNewAbsCurv
        'Modif de l'info-bulle
        'unRep.monIcone.ToolTipText = unRep.monNomCourt + " / Type : " + unRep.monIcone.Tag + " / AbsCurv = " + Format(unRep.monAbsCurv) + " m"
        unRep.monIcone.ToolTipText = unRep.monNomCourt + " / Type : " + DonnerIconeRepere(unRep.monTypeIcone).Tag + " / AbsCurv = " + Format(unRep.monAbsCurv) + " m"
        
        'Stockage des anciens min et max en distance
        unOldMinD = uneFrmD.monMinD
        unOldMaxD = uneFrmD.monMaxD
        
        'On n'est jamais en section de travail pour les modifs
        'des abscisses curvilignes
        If (unNewAbsCurv > uneFrmD.monMaxD Or unAncienAbsCurvSup = uneFrmD.monMaxD) And unAncienAbsCurvInf > uneFrmD.monMinD Then
            'Cas où l'abs curv modifié est > au maxi en distance
            'ou était égale avant à ce maxi mais sans être le mini avant
            '==> on redessine tout au bon niveau de zoom
            unYRepMax = DonnerYRepMax(uneFrmD)
            uneFrmD.monMaxD = unYRepMax
            If unNewAbsCurv < uneFrmD.monMinD Then
                'Cas où l'ancien max devient le min
                uneFrmD.monMinD = unNewAbsCurv
                uneFrmD.monMinDtot = unNewAbsCurv
            End If
            If unOldMinD <> uneFrmD.monMinD Or unOldMaxD <> uneFrmD.monMaxD Then
                'Si vrai changement d'englobant min et max en distance
                RedessinerZoomTout uneFrmD
            End If
            'Recalcul de la longueur du parcours de référence
            uneFrmD.monMaxDtot = unYRepMax
            'uneFrmD.TextLongIti.Text = Format(uneFrmD.monMaxDtot - uneFrmD.monMinDtot)
            uneFrmD.TextLongIti.Text = Format(DonnerLongIti(uneFrmD))
        ElseIf (unNewAbsCurv < uneFrmD.monMinD Or unAncienAbsCurvInf = uneFrmD.monMinD) And unAncienAbsCurvSup < uneFrmD.monMaxD Then
            'Cas où l'abs curv modifié est < au mini en distance
            'ou était égale avant à ce mini mais sans être le maxi avant
            '==> on redessine tout au bon niveau de zoom
            unYRepMin = DonnerYRepMin(uneFrmD)
            uneFrmD.monMinD = unYRepMin
            If unNewAbsCurv > uneFrmD.monMaxD Then
                'Cas où l'ancien min devient le max
                uneFrmD.monMaxD = unNewAbsCurv
                uneFrmD.monMaxDtot = unNewAbsCurv
            End If
            If unOldMinD <> uneFrmD.monMinD Or unOldMaxD <> uneFrmD.monMaxD Then
                'Si vrai changement d'englobant min et max en distance
                RedessinerZoomTout uneFrmD
            End If
            'Recalcul de la longueur du parcours de référence
            uneFrmD.monMinDtot = unYRepMin
            'uneFrmD.TextLongIti.Text = Format(uneFrmD.monMaxDtot - uneFrmD.monMinDtot)
            uneFrmD.TextLongIti.Text = Format(DonnerLongIti(uneFrmD))
        Else
            DessinerRepere uneFrmD, unRep, unAncienAbsCurv
        End If
    End If
End Sub


Public Sub RemplirSpreadParcours(uneFrmD As frmDocument, Optional uneMAJParMoyen As Boolean = False)
    'Remplissage du tableau des parcours affectés d'une form itinéraire
    Dim unPar As Parcours, unNbPar As Integer
    
    uneFrmD.SpreadParcours.MaxRows = uneFrmD.maColParcours.Count
    
    If uneMAJParMoyen Then
        'Cas de la mise à jour uniquement de la ligne 1,
        'celle du parcours moyen
        unNbPar = 1
    Else
        'Cas de la mise à jour de tous les lignes, donc de tous les parcours
        unNbPar = uneFrmD.maColParcours.Count
    End If
    
    For i = 1 To unNbPar
        Set unPar = uneFrmD.maColParcours(i)
        uneFrmD.SpreadParcours.Row = i
        uneFrmD.SpreadParcours.Col = 1
        uneFrmD.SpreadParcours.Text = unPar.monNom
        If uneMAJParMoyen = False Then
            uneFrmD.SpreadParcours.Col = 2
            uneFrmD.SpreadParcours.Value = Abs(unPar.monIsUtil)
        End If
        uneFrmD.SpreadParcours.Col = 3
        uneFrmD.SpreadParcours.BackColor = unPar.maCouleur
        uneFrmD.SpreadParcours.Col = 4
        uneFrmD.SpreadParcours.Text = unPar.monEnqueteur
        uneFrmD.SpreadParcours.Col = 5
        uneFrmD.SpreadParcours.Text = unPar.monNumVeh
        uneFrmD.SpreadParcours.Col = 6
        uneFrmD.SpreadParcours.TypeComboBoxCurSel = unPar.maMeteo
        uneFrmD.SpreadParcours.Col = 7
        uneFrmD.SpreadParcours.Text = unPar.maDate
        uneFrmD.SpreadParcours.Col = 8
        uneFrmD.SpreadParcours.Text = unPar.monJourSemaine
        uneFrmD.SpreadParcours.Col = 9
        uneFrmD.SpreadParcours.Text = unPar.monHeureDebut
        uneFrmD.SpreadParcours.Col = 10
        uneFrmD.SpreadParcours.Text = UBound(unPar.monTabAbsRep)
        uneFrmD.SpreadParcours.Col = 11
        'On arrondit au mètre près la distance parcourue
        uneFrmD.SpreadParcours.Text = CLng(unPar.maDistPar * unPar.monCoefEta / 10)
        uneFrmD.SpreadParcours.Col = 12
        'On arrondit au mètre la distance parcourue au dernier top
        If DonnerNbParcoursUtil(uneFrmD) > 0 Then
            'Cas où il y des parcours utilisés
            uneFrmD.SpreadParcours.Text = CLng(unPar.monTabAbsRep(UBound(unPar.monTabAbsRep)) * unPar.monCoefEta / 10)
        Else
            'Cas où aucun parcours n'est utilisé
            uneFrmD.SpreadParcours.Text = 0
        End If
        
        uneFrmD.SpreadParcours.Col = 13
        
        'Formattage en 00h 00mn 00s de la durée
        uneStringDuree = FormatterTempsEnHMNS(unPar.maDuree)
        uneFrmD.SpreadParcours.Text = uneStringDuree
        
        uneFrmD.SpreadParcours.Col = 14
        uneFrmD.SpreadParcours.Text = unPar.monCoefEta
    Next i
End Sub

Public Sub RemplirMeteoSpreadParcours(uneFrmD As frmDocument)
    'Remplissage du spread parcours avec les libellés météo de la form itinéraire
    For i = 1 To uneFrmD.SpreadParcours.MaxRows
        uneFrmD.SpreadParcours.Row = i
        uneFrmD.SpreadParcours.Col = 6 'Condition météo en colonne 6
        'Stockage de l'item sélectionné dans la cellule combobox
        unIndSel = uneFrmD.SpreadParcours.TypeComboBoxCurSel
        'On vide la cellule de type combobox
        uneFrmD.SpreadParcours.Action = 26 ' = SS_ACTION_COMBO_CLEAR
        For j = 1 To uneFrmD.maColMeteo.Count
            'Remplissage de la combox de la cellule ligne i et colonne 6
            uneFrmD.SpreadParcours.TypeComboBoxIndex = j - 1
            uneFrmD.SpreadParcours.TypeComboBoxString = uneFrmD.maColMeteo(j)
            'Restauration de l'item sélectionné dans la cellule combobox
            uneFrmD.SpreadParcours.TypeComboBoxCurSel = unIndSel
        Next j
    Next i
End Sub

Public Function DonnerLongIti(uneForm As Form) As Long
    'Fonction donnant la longueur de l'itinéraire total
    'ou de la section de travail
    Dim unRep1 As Repere, unRep2 As Repere
    
    If uneForm.CheckSection.Value = 0 Then
        'Pas de section de travail définie
        DonnerLongIti = CLng(uneForm.monMaxDtot - uneForm.monMinDtot)
    Else
        'Cas d'une section de travail définie
        Set unRep1 = uneForm.maColRepere(uneForm.ComboRepFinSec.ListIndex + 1)
        Set unRep2 = uneForm.maColRepere(uneForm.ComboRepDebSec.ListIndex + 1)
        DonnerLongIti = Abs(unRep1.monAbsCurv - unRep2.monAbsCurv)
    End If
End Function

Public Sub CocherToutColSelection(unTabSpread As vaSpread)
    'Procédure cochant toutes les cases de la colonne
    'Sélection ( = colonne n°2 dans MiTemps) d'un spread
    For i = 1 To unTabSpread.MaxRows
        unTabSpread.Row = i
        unTabSpread.Col = 2
        unTabSpread.Value = 1 'case cochée
    Next i
End Sub

Public Sub DecocherToutColSelection(unTabSpread As vaSpread)
    'Procédure décochant toutes les cases de la colonne
    'Sélection ( = colonne n°2 dans MiTemps) d'un spread
    For i = 1 To unTabSpread.MaxRows
        unTabSpread.Row = i
        unTabSpread.Col = 2
        unTabSpread.Value = 0 'case décochée
    Next i
End Sub

