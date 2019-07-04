Attribute VB_Name = "ModuleMain"
Declare Function GetDiskFreeSpace Lib "Kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

'Variables globales
'Public formMain As frmMain
Public mesOptions As New Options
Public maPicBox As PictureBox

'Création du parcours à importer qui sert à remplir un fichier MIT
Public monParToImport As New Parcours

'Collection stockant les abs curvilignes moyennes des repères des parcours
'de la collection maColAbsRep avec en plus en dernière place la distance
'parcourue moyenne de ces parcours
Public maColValMoy As New Collection

'Variable permettant de savoir quel bouton d'une fenêtre de dialogue
'a été cliqué avant sa fermeture
Public monBtnClick As String

'Variable stockant la liste des parcours stockés dans le mtb lu
Public maColParcoursMTB As New ColParcours

'Constantes pour la taille initiale des spread repères et parcours
'avant tout retaillage (hauteur et top)
Public Const HSpreadRep As Integer = 2295
Public Const HSpreadPar As Integer = 2055
Public Const TSpreadRep As Integer = 1440
Public Const TSpreadPar As Integer = 4320 '3840

'Constantes pour l'épaisseur des traits des courbes DV et DT
Public Const TraitFin As Byte = 1
Public Const TraitGros As Byte = 4

'Constantes pour la marge d'erreur en visibilité à l'écran
Public Const EpsilonEcran As Integer = 100 ' en Twips

'Constante pour l'epsilon servant aux tests de proximité en distance
'On affiche des mètres mais les mesures de distances sont en décimètres,
'On prend une précision de 1.1 dm = 11 cm (avant bug CEP 0.01 = 0.01 dm = 0.1 cm)
Public Const EpsilonDist As Single = 1.1 '0.01

'Constante de précision pour les tests d'égalité entre flottant
Public Const Epsilon As Single = 0.001

'Constantes pour le type de format de fichier des fichiers MTB
Public Const OldMTB As Byte = 1     'vieux format
Public Const NewMTB As Byte = 2     'nouveau format
Public Const BadMTB As Byte = 3     'mauvais format

'Constantes pour la taille du tableau des pas d'inter-distances
'NbPasMax = 86400 pas = 86400 sec si le pas de mesure = le mini possible valant 1s
'Donc 86400s / 3600s = 24 heures de mesures avec un pas de 1 seconde = le pas mini
Public Const NbPasMax As Long = 86400

'Constantes pour les types d'icones de repères
'Classés dans le même ordre que les items de la combox
'de la colonne Type de repère du spread des repères
Public Const Feu As Byte = 1
Public Const PanneauStop As Byte = 2
Public Const CedezPassage As Byte = 3
Public Const Carrefour As Byte = 4
Public Const Giratoire As Byte = 5
Public Const EntreeAgglo As Byte = 6
Public Const SortieAgglo As Byte = 7
Public Const ArretBus As Byte = 8
Public Const PassagePieton As Byte = 9
Public Const Peage As Byte = 10
Public Const EntreeAuto As Byte = 11
Public Const SortieAuto As Byte = 12
Public Const StationService As Byte = 13
Public Const AireRepos As Byte = 14
Public Const Autre As Byte = 15
Public Const DoubleTop As Byte = 16

'Constante pour la taille des repères de la
'fenêtre importation campagne mesures
Public Const TailleRep As Byte = 60

'Constante pour un cm = 567 twips
Public Const UnCmEnTwips As Integer = 567
Public Const CarreDe256 As Long = 65536

'Constante pour la position Top et les marges haut et bas
'des pictures box des onglets
Public Const PicBoxTop As Integer = 60
Public Const PicBoxMargeH As Integer = 150  '195 = Hauteur en twips du label autosizé valant 999999
Public Const PicBoxMargeB As Integer = 300  '195 = Hauteur en twips du label autosizé valant 9999

'Constantes pour le numéro d'onglet d'une fenêtre fille
Public Const OngletItiRef As Integer = 0
Public Const OngletCbeDT As Integer = 1
Public Const OngletCbeDV As Integer = 2
Public Const OngletSynoV As Integer = 3
Public Const OngletHistV As Integer = 4
Public Const OngletTabBr As Integer = 5
Public Const OngletTabSS As Integer = 6

'Paramétrages de l'aide contextuel
'Constante pour définir le nom du fichier d'aide
Public Const Help_Chm As String = "\MiTemps.chm"
'Constante pour le nom de la fenêtre principale de l'aide
Public Const Help_Main As String = "Main"

'Constantes pour les numéro de HelpContextID
Public Const HelpID_WinNew As Integer = 205
Public Const HelpID_WinOpen As Integer = 207
Public Const HelpID_WinSave As Integer = 210
Public Const HelpID_WinSaveAs As Integer = 211
Public Const HelpID_WinClose As Integer = 208
Public Const HelpID_WinRabouter As Integer = 209
Public Const HelpID_WinNewByMesure As Integer = 206
Public Const HelpID_WinImportMesure As Integer = 212
Public Const HelpID_WinViderBoitier As Integer = 213
Public Const HelpID_WinPrint As Integer = 214
Public Const HelpID_WinQuit As Integer = 215
Public Const HelpID_WinOngletItiRef As Integer = 217
Public Const HelpID_WinOngletCbeDT As Integer = 218
Public Const HelpID_WinOngletCbeDV As Integer = 219
Public Const HelpID_WinOngletSynoV As Integer = 220
Public Const HelpID_WinOngletHistoV As Integer = 221
Public Const HelpID_WinOngletTabBrut As Integer = 222
Public Const HelpID_WinOngletTabStat As Integer = 223
Public Const HelpID_WinBarOutil As Integer = 225
Public Const HelpID_WinBarEtat As Integer = 226
Public Const HelpID_WinOptions As Integer = 227
Public Const HelpID_WinAPropos As Integer = 236

'Variable globale pointant sur l'itinéraire en cours (fenêtre fille active)
Public monIti As Form

'Variable donnant l'index dans le groupe de controle ImageRepere
'du dernier cliqué, donc de celui sélectionné
Public monIndIconeRepSel As Integer


Sub Main()
    'Initialisation du printer pour MiTemps
    If Printers.Count > 0 Then InitialiserPrinter
    
    '********************************
    'test QLM module Qlm
    '********************************
  'Type de protection
        TYPPROTECTION = CPM
      ' Vérification de l'enregistrement
      If ProtectCheck("its00+-k") = "its00+-k" Then
        ' Affichage de la feuille principale
         frmMain.Show
      Else 'la licence n'a pas été validée on ferme
         End
      End If
    '********************************
    
    'Récupération des options dans la base de registre
    RécupérerOptions
        
    'Création de la MDI mère
    'Set formMain = New frmMain
    'formMain.Show
    
End Sub

Public Sub RécupérerOptions()
    'Récupération des options générales par lecture des valeurs
    'de ces options stockées dans la base de registre
    Dim uneStrTmp As String
    
    mesOptions.maTolLong = GetSetting(App.Title, "Options", "ToléranceLongueur", 5)
    mesOptions.monEcartMax = GetSetting(App.Title, "Options", "EcartMax", 1)
    
    mesOptions.maCouleurClasV1 = GetSetting(App.Title, "Options", "ColorClasV1", QBColor(12))
    mesOptions.maValClasV1 = GetSetting(App.Title, "Options", "ValClasV1", 30)
    mesOptions.maCouleurClasV2 = GetSetting(App.Title, "Options", "ColorClasV2", QBColor(13))
    mesOptions.maValClasV2 = GetSetting(App.Title, "Options", "ValClasV2", 50)
    mesOptions.maCouleurClasV3 = GetSetting(App.Title, "Options", "ColorClasV3", QBColor(14))
    mesOptions.maValClasV3 = GetSetting(App.Title, "Options", "ValClasV3", 70)
    mesOptions.maCouleurClasV4 = GetSetting(App.Title, "Options", "ColorClasV4", QBColor(10))
    mesOptions.maValClasV4 = GetSetting(App.Title, "Options", "ValClasV4", 90)
    mesOptions.maCouleurClasV5 = GetSetting(App.Title, "Options", "ColorClasV5", QBColor(11))
    mesOptions.maValClasV5 = GetSetting(App.Title, "Options", "ValClasV5", 110)
    mesOptions.maCouleurClasV6 = GetSetting(App.Title, "Options", "ColorClasV6", QBColor(9))

    'Remplissage des libellés par défaut des conditions météo
    uneStrTmp = "Indéfinie"
    uneStrTmp = uneStrTmp + ",Beau Temps"
    uneStrTmp = uneStrTmp + ",Pluie Forte"
    uneStrTmp = uneStrTmp + ",Pluie Légère"
    uneStrTmp = uneStrTmp + ",Neige"
    uneStrTmp = uneStrTmp + ",Grêle"
    uneStrTmp = uneStrTmp + ",Brouillard"
    uneStrTmp = uneStrTmp + ",Vent Fort"
    uneStrTmp = uneStrTmp + ",Tempête"
    mesOptions.mesLibMeteo = GetSetting(App.Title, "Options", "LibMeteo", uneStrTmp)
    
    'Epaisseur des traits pour les dessins de courbes
    mesOptions.monEpaisTrait = GetSetting(App.Title, "Options", "EpaisTrait", 1)
End Sub

Public Sub StockerOptions()
    'Stockage des options dans la base de registre
    With mesOptions
        SaveSetting App.Title, "Options", "ToléranceLongueur", .maTolLong
        SaveSetting App.Title, "Options", "EcartMax", .monEcartMax
        
        SaveSetting App.Title, "Options", "ColorClasV1", .maCouleurClasV1
        SaveSetting App.Title, "Options", "ValClasV1", .maValClasV1
        SaveSetting App.Title, "Options", "ColorClasV2", .maCouleurClasV2
        SaveSetting App.Title, "Options", "ValClasV2", .maValClasV2
        SaveSetting App.Title, "Options", "ColorClasV3", .maCouleurClasV3
        SaveSetting App.Title, "Options", "ValClasV3", .maValClasV3
        SaveSetting App.Title, "Options", "ColorClasV4", .maCouleurClasV4
        SaveSetting App.Title, "Options", "ValClasV4", .maValClasV4
        SaveSetting App.Title, "Options", "ColorClasV5", .maCouleurClasV5
        SaveSetting App.Title, "Options", "ValClasV5", .maValClasV5
        SaveSetting App.Title, "Options", "ColorClasV6", .maCouleurClasV6
    
        SaveSetting App.Title, "Options", "LibMeteo", .mesLibMeteo
    End With
End Sub

Public Sub ChoisirCouleur(unePicCouleur As PictureBox)
    'Choix de la couleur parmi les couleurs systèmes disponibles
    'pour la PictureBox passée en paramètre
    With frmMain
          ' Attribue à CancelError la valeur True
          .dlgCommonDialog.CancelError = True
          On Error GoTo ErrHandler
          ' Définit la propriété Flags
          .dlgCommonDialog.flags = cdlCCRGBInit
          ' Affiche la boîte de dialogue Couleur
          .dlgCommonDialog.ShowColor
          ' Attribue à l'arrière-plan de la feuille la
          ' couleur sélectionnée
          unePicCouleur.BackColor = .dlgCommonDialog.Color
    End With
      
    Exit Sub

ErrHandler:
    ' L'utilisateur a cliqué sur Annuler
    'On ne fait rien
End Sub


Public Sub DessinerCourbe(uneForm As Form, uneZoneDessin As Object, unTypeCourbe As Integer)
    'Dessin d'une courbe DT ou DV (cf valeur de unTypeCourbe) sur une zone
    'de dessin écran ou imprimante d'une étude (= uneForm)
    '(écran = picture box de l'onglet courbeDT)
    Dim uneMargeG As Single, uneMargeD As Single 'marges gauche et droite
    Dim uneMargeH As Single, uneMargeB As Single 'marges haut et bas
    Dim unLibX As String, uneHText As Single, uneWText As Single
    Dim unMaxXecran As Single, unMaxYecran As Single
    Dim unMinXecran As Single, unMinYecran As Single
    Dim unMaxXreel As Single, unMinXreel As Single
    Dim unMaxXreelTmp As Single, unMinXreelTmp As Single
    Dim unMaxYreel As Single, unMinYreel As Single
    Dim unXecran As Single, unYecran As Single, unXFinDessin As Single
    Dim unXecSuiv As Single, unXecPred As Single
    Dim unYecSuiv As Single, unYecPred As Single
    Dim unPasXGrad1Tmp As Long, unPasXGrad2Tmp As Long
    Dim unPasXGrad1 As Single, unPasXGrad2 As Single
    Dim unPasYGrad1 As Long, unPasYGrad2 As Long, unMaxWidthNomRep As Single
    Dim j As Long, unNbPoints As Long, unRep As Repere
    Dim uneDistMaxReelX As Single, uneDistMaxEcranX As Single
    Dim uneDistMaxReelY As Single, uneDistMaxEcranY As Single
    Dim unRapX As Single, unRapY As Single, unParcours As Parcours
    
    'Affichage du sablier en pointeur souris pour symboliser l'attente
    uneForm.MousePointer = vbHourglass
    unPrint = (TypeOf uneZoneDessin Is Printer) '=-1 si impression, 0 sinon
    
    'Initialisation de variables utilisées ci-après
    uneZoneDessin.Font.Name = "Arial"
    uneZoneDessin.Font.Size = 7
    uneHText = uneZoneDessin.TextHeight("9")
    
    'On fixe les tailles de marges
    If TypeOf uneZoneDessin Is Printer Then
        unMaxWidthNomRep = FixerMargesImprimante(uneMargeG, uneMargeD, uneMargeH, uneMargeB)
    ElseIf TypeOf uneZoneDessin Is PictureBox Then
        uneZoneDessin.Cls
        FixerMargesPicBox uneForm, uneZoneDessin, uneMargeG, uneMargeD, uneMargeH, uneMargeB
    Else
        MsgBox MsgErreurProg + MsgErreurTypeZoneDessinInconnu + MsgIn + "ModuleMain:DessinerCourbe", vbCritical
        Exit Sub
    End If
    
    'Stockage des y min et max réels
    unMaxYreel = uneForm.monMaxD
    unMinYreel = uneForm.monMinD
    
    'Calcul des pas de graduations primaires et secondaires et arrondis
    'de la valeur mini à la graduation secondaire juste inférieure
    'et de la valeur maxi à la graduation secondaire juste supérieure
    TrouverPasGradEtModifierMinMax unPasYGrad1, unPasYGrad2, unMaxYreel, unMinYreel
    
    'Stockage des y min et max réels en distances si on m'imprime pas
    If unPrint = False Then
        uneForm.monMaxD = unMaxYreel
        uneForm.monMinD = unMinYreel
        'Stockage des pas de graduations en distance
        uneForm.monPasGrad1 = unPasYGrad1
        uneForm.monPasGrad2 = unPasYGrad2
    End If
    
    If unTypeCourbe = OngletCbeDT Then
        'Dessin d'une courbe distance-temps dans la zone de dessin
        'Affichage du libellé de l'axe des X
        unLibX = "T (mn)"
        'Stockage des max et min x réels
        unMaxXreel = uneForm.monMaxT
        unMinXreel = uneForm.monMinT
    ElseIf unTypeCourbe = OngletCbeDV Then
        'Dessin d'une courbe distance-vitesse dans la zone de dessin
        'Affichage du libellé de l'axe des X
        unLibX = "V (km/h)"
        'Stockage des max et min x réels
        unMaxXreel = uneForm.monMaxV
        unMinXreel = uneForm.monMinV
    Else
        MsgBox MsgErreurProg + MsgErreurTypeCourbeInconnu + MsgIn + "ModuleMain:DessinerCourbe", vbCritical
        Exit Sub
    End If
    
    'Calcul des pas de graduations primaires et secondaires et arrondis
    'de la valeur mini à la graduation secondaire juste inférieure
    'et de la valeur maxi à la graduation secondaire juste supérieure
    ArrondirXMinXMaxGrad2 unPasXGrad1, unPasXGrad2, unMaxXreel, unMinXreel
    
    If unTypeCourbe = OngletCbeDT And unPrint = False Then
        'Stockage des max et min x réels si on m'imprime pas
        uneForm.monMaxT = unMaxXreel
        uneForm.monMinT = unMinXreel
        'Remplissage des info sur le parcours sélectionné dans l'onglet courbe DT
        RemplirSpreadInfoParcoursSel uneForm.SpreadInfoParcoursDT, uneForm, uneForm.monIndParcoursSelectDT
    ElseIf unTypeCourbe = OngletCbeDV And unPrint = False Then
        'Stockage des max et min x réels si on m'imprime pas
        uneForm.monMaxV = unMaxXreel
        uneForm.monMinV = unMinXreel
        'Remplissage des info sur le parcours sélectionné dans l'onglet courbe DV
        RemplirSpreadInfoParcoursSel uneForm.SpreadInfoParcoursDV, uneForm, uneForm.monIndParcoursSelectDV
    End If
    
    'Variables servant pour la conversion coordonnées réelles en écran
    uneDistMaxReelX = unMaxXreel - unMinXreel
    uneDistMaxEcranX = uneZoneDessin.Width - uneMargeG - uneMargeD
    uneDistMaxReelY = unMaxYreel - unMinYreel
    uneDistMaxEcranY = uneZoneDessin.Height - uneMargeH - uneMargeB

    'Conversion en coordonnées Y écran des distances réelles
    unMaxYecran = uneMargeH
    unMinYecran = ConvertirEnEcran(unMaxYecran, unMaxYreel - unMinYreel, unMaxYreel - unMinYreel, uneZoneDessin.Height - uneMargeH - uneMargeB)
    
    'Conversion en coordonnées X écran des temps ou vitesses réels
    unMinXecran = uneMargeG
    unMaxXecran = ConvertirEnEcran(unMinXecran, unMaxXreel - unMinXreel, unMaxXreel - unMinXreel, uneZoneDessin.Width - uneMargeG - uneMargeD)
    
    'Stockage dans la form du minimum écran en X et du maximum écran en Y
    'utilisé par la fonction SelectionnerParcours si on m'imprime pas
    If unPrint = False Then
        uneForm.monMaxYecran = unMaxYecran
        uneForm.monMinXecran = unMinXecran
    End If
    
    'Dessin en trait continu fin
    uneZoneDessin.DrawWidth = TraitFin
    uneZoneDessin.DrawStyle = vbSolid
    
    'Mise des graduations sur OY
    unAfficheToutY = ((unMaxYreel - unMinYreel) < 26 * unPasYGrad2)
    For i = unMinYreel To unMaxYreel Step unPasYGrad2
        'Dessin du trait de graduation primaire ou secondaire
        unYecran = ConvertirEnEcran(unMaxYecran, unMaxYreel - i, unMaxYreel - unMinYreel, uneZoneDessin.Height - uneMargeH - uneMargeB)
        uneZoneDessin.Line (unMinXecran, unYecran)-(unMinXecran - 60 * (1 - (i Mod unPasYGrad1 = 0)), unYecran), 0

        'Placement en X écran
        uneZoneDessin.CurrentX = unMinXecran - uneZoneDessin.TextWidth(Format(i)) - 120
        
        If i Mod unPasYGrad1 = 0 Or unAfficheToutY Then
            'Affichage de la valeur de la graduation primaire
            'ou Affichage des valeurs des graduations secondaires si au plus
            '26 sont à afficher, contrainte de 26 due aux écrans 800x600
            uneZoneDessin.CurrentY = unYecran - uneZoneDessin.TextHeight(Format(i)) / 2
            uneZoneDessin.Print Format(i)
        ElseIf i = unMaxYreel Then
            'Affichage de la valeur max
            'Si le max est une graduation primaire ou secondaire,
            'sa position d'affichage est gérée par le 1er if ci-dessus.
            'Sinon l'affichage du max est décalée vers le haut
            'pour éviter les superpositions avec les graduations
            'Calcul du Y écran de la graduation précédente
            unYecPred = ConvertirEnEcran(unMaxYecran, unMaxYreel - i + unPasYGrad2, unMaxYreel - unMinYreel, uneZoneDessin.Height - uneMargeH - uneMargeB)
            'Affichage de la valeur max avec ajout d'un décalage si impression
            unDecalage = Abs(unYecPred - unYecran < uneHText) '=1 si vrai, 0 si faux
            uneZoneDessin.CurrentY = (unYecran - uneHText / 2) * (1 - unDecalage) + (unYecPred - uneHText * (1.2 - unPrint * 0.2)) * unDecalage
            uneZoneDessin.Print Format(i)
        ElseIf i = unMinYreel Then
            'Affichage de la valeur min
            'Si le min est une graduation primaire ou secondaire,
            'sa position d'affichage est gérée par le 1er if ci-dessus.
            'Sinon l'affichage du min est décalée vers le bas.
            'pour éviter les superpositions avec les graduations
            'Calcul du Y écran de la graduation suivante
            unYecSuiv = ConvertirEnEcran(unMaxYecran, unMaxYreel - i - unPasYGrad2, unMaxYreel - unMinYreel, uneZoneDessin.Height - uneMargeH - uneMargeB)
            'Affichage de la valeur min avec ajout d'un décalage si impression
            unDecalage = Abs(unYecran - unYecSuiv < uneHText) '=1 si vrai, 0 si faux
            uneZoneDessin.CurrentY = (unYecran - uneHText / 2) * (1 - unDecalage) + (unYecSuiv + uneHText * (0.2 - unPrint * 0.2)) * unDecalage
            uneZoneDessin.Print Format(i)
        End If
    Next i
    
    'Affichage de la valeur et son unité affichée sur Oy
    If unPrint Then
        uneZoneDessin.CurrentX = unMinXecran - uneZoneDessin.TextWidth("Distance (m)") / 2
        uneZoneDessin.CurrentY = unMaxYecran - uneHText * 2
        uneZoneDessin.Print "Distance (m)"
    End If
    
    'Mise des graduations sur OX
    'Code de parcours de boucle différent car on peut avoir des sous-divisions à
    'un chiffre prés la virgule
    unAfficheToutX = ((unMaxXreel - unMinXreel) < 26 * unPasXGrad2)
    i = unMinXreel
    Do
        i = CSng(Format(i, "0#####.#"))
        unXecran = ConvertirEnEcran(unMinXecran, i - unMinXreel, unMaxXreel - unMinXreel, uneZoneDessin.Width - uneMargeG - uneMargeD)
        
        'Placement en Y écran décalée plus vers le bas si on imprime
        'pour que les graduations décalées à droite et à gauche des min et max
        'sur OX ne chevauchent pas les graduations de OY
        unCurY = unMinYecran + uneZoneDessin.TextHeight(Format(i)) / 2 * (1 - unPrint)
        uneZoneDessin.CurrentY = unCurY
        
        If (i * 10) Mod (unPasXGrad1 * 10) = 0 Or unAfficheToutX Then
            'Affichage de la valeur de la graduation primaire
            'ou Affichage des valeurs des graduations secondaires si au plus
            '26 sont à afficher, contrainte de 26 due aux écrans 800x600
            unCurX = unXecran - uneZoneDessin.TextWidth(Format(i)) / 2
            uneZoneDessin.CurrentX = unCurX
            uneZoneDessin.Print Format(i)
        ElseIf Abs(i - unMaxXreel) < Epsilon Then
            'Affichage de la valeurs max
            'Si le max est une graduation primaire ou secondaire,
            'sa position d'affichage est gérée par le if ci-dessus.
            'Sinon pour l'affichage du max est décalée vers la droite
            'pour éviter les superpositions avec les graduations
            'Calcul du X écran de la graduation précédent
            unXecPred = ConvertirEnEcran(unMinXecran, i - unPasXGrad2 - unMinXreel, unMaxXreel - unMinXreel, uneZoneDessin.Width - uneMargeG - uneMargeD)
            'Affichage de la valeur max
            uneWText = uneZoneDessin.TextWidth(Format(i))
            unDecalage = Abs(unXecran - unXecPred < uneWText) '=1 si vrai, 0 si faux
            unCurX = (unXecran - uneWText / 2) * (1 - unDecalage) + (unXecPred + uneWText * 0.65) * unDecalage
            uneZoneDessin.CurrentX = unCurX
            uneZoneDessin.Print Format(i)
            'Création d'une ligne reliant le max avec son affichage décalé
            If unDecalage = 1 Then uneZoneDessin.Line (unCurX + uneWText * 0.15, unCurY)-(unXecran, unMinYecran + 60), 0
        ElseIf Abs(i - unMinXreel) < Epsilon Then
            'Affichage de la valeurs min
            'Si le min est une graduation primaire ou secondaire,
            'sa position d'affichage est gérée par le if ci-dessus.
            'Sinon pour l'affichage du min est décalée vers la gauche
            'pour éviter les superpositions avec les graduations
            'Calcul du X écran de la graduation précédent
            unXecSuiv = ConvertirEnEcran(unMinXecran, i + unPasXGrad2 - unMinXreel, unMaxXreel - unMinXreel, uneZoneDessin.Width - uneMargeG - uneMargeD)
            'Affichage de la valeur min
            uneWText = uneZoneDessin.TextWidth(Format(i))
            unDecalage = Abs(unXecSuiv - unXecran < uneWText) '=1 si vrai, 0 si faux
            unCurX = (unXecran - uneWText / 2) * (1 - unDecalage) + (unXecSuiv - uneWText * 1.75) * unDecalage
            uneZoneDessin.CurrentX = unCurX
            uneZoneDessin.Print Format(i)
            'Création d'une ligne reliant le min avec son affichage décalé
            If unDecalage = 1 Then uneZoneDessin.Line (unCurX + uneWText * 0.85, unCurY)-(unXecran, unMinYecran + 60), 0
        End If
        
        'Dessin du trait de graduation primaire ou secondaire en dernier
        'pour qu'il ne soit pas caché par le texte de la valeur de graduation
        'Surtout en impression papier
        uneZoneDessin.Line (unXecran, unMinYecran)-(unXecran, unMinYecran + 60 * (1 - (i Mod unPasXGrad1 = 0) * (0.66 - 0.34 * unPrint))), 0
        
        i = i + unPasXGrad2
    Loop While i <= unMaxXreel
    
    'Dessiner la courbe du bon type (DT ou DV) de chacun des parcours
    For i = 1 To uneForm.maColParcours.Count
        Set unParcours = uneForm.maColParcours(i)
        If unParcours.monIsUtil Then
            If unPrint Then
                uneZoneDessin.DrawWidth = mesOptions.monEpaisTrait
            Else
                uneZoneDessin.DrawWidth = TraitFin
            End If
            uneCouleur = unParcours.maCouleur
            unNbPoints = unParcours.monNbPas
            'Récup des données pour la courbe temps/distance ou vitesse/distance
            If unTypeCourbe = OngletCbeDT Then
                unD1 = 0
                unX1 = 0
                'Dessin en plus gros si c'est le parcours sélectionné
                If uneForm.monIndParcoursSelectDT = i And unPrint = False Then uneZoneDessin.DrawWidth = TraitGros
                'Conversion des distances des décimètres au mètre
                unD2 = unParcours.monTabDist(1) / 10 * unParcours.monCoefEta
                'Conversion des temps des dixièmes de seconde au minute
                unX2 = unParcours.monFirstPas / 600
            ElseIf unTypeCourbe = OngletCbeDV Then
                unD1 = 0
                unX1 = 0
                'Dessin en plus gros si c'est le parcours sélectionné
                If uneForm.monIndParcoursSelectDV = i And unPrint = False Then uneZoneDessin.DrawWidth = TraitGros
                'Conversion des distances des décimètres au mètre
                unD2 = unParcours.monTabDist(1) / 10 * unParcours.monCoefEta
                If unParcours.monFirstPas = 0 Then
                    unX2 = 0
                Else
                    'mètre/dixième de seconde converti en km/h
                    unX2 = unD2 / unParcours.monFirstPas * 36
                End If
            Else
                MsgBox MsgErreurProg + MsgErreurTypeCourbeInconnu + MsgIn + "ModuleMain:DessinerCourbe", vbCritical
                Exit Sub
            End If
            
            'Conversion en coordonnées écrans des coordonnées réelles
            'du premier point =(unX1, unD1) et du deuxième point =(unX2, unD2)
            unXecran = ConvertirEnEcran(unMinXecran, unX1 - unMinXreel, uneDistMaxReelX, uneDistMaxEcranX)
            unYecran = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD1, uneDistMaxReelY, uneDistMaxEcranY)
            unXecSuiv = ConvertirEnEcran(unMinXecran, unX2 - unMinXreel, uneDistMaxReelX, uneDistMaxEcranX)
            unYecSuiv = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD2, uneDistMaxReelY, uneDistMaxEcranY)
            'Dessin de la courbe du premier segment
            's'il est entre le min et le max y écran
            'min y écran  > max y écran car les y écran orientés vers le bas en Y,
            'donc aprés conversion donnée réelle en écran le max devient < au min
            'If unYecran <= unMinYecran And unYecran >= unMaxYecran Then
            If (unYecran <= unMinYecran And unYecran >= unMaxYecran) Or (unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran) Then
                uneZoneDessin.Line (unXecran, unYecran)-(unXecSuiv, unYecSuiv), uneCouleur
            End If
            'Stockage pour le segment suivant
            unXecran = unXecSuiv
            unYecran = unYecSuiv
            
            For j = 2 To unNbPoints - 1
                'Calcul du point suivant pour la courbe temps/distance
                'ou la courbe vitesse/distance
                'Cumul des distances et Conversion des distances
                'des décimètres au mètre
                unD2 = unD2 + unParcours.monTabDist(j) / 10 * unParcours.monCoefEta
                If unTypeCourbe = OngletCbeDT Then
                    'Cumul des temps et Conversion du pas des secondes en minute
                    unX2 = unX2 + unParcours.monPasMesure / 60
                ElseIf unTypeCourbe = OngletCbeDV Then
                    'Décimètre/seconde converti en km/h
                    unX2 = unParcours.monTabDist(j) * unParcours.monCoefEta / unParcours.monPasMesure * 0.36
                End If
                
                'Conversion en coordonnées écrans des coordonnées réelles
                'du point suivant
                unXecSuiv = ConvertirEnEcran(unMinXecran, unX2 - unMinXreel, uneDistMaxReelX, uneDistMaxEcranX)
                unYecSuiv = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD2, uneDistMaxReelY, uneDistMaxEcranY)
                'Dessin de la courbe segment par segment
                's'il est entre le min et le max y écran
                'min y écran  > max y écran car les y écran orientés vers le bas en Y,
                'donc aprés conversion donnée réelle en écran le max devient < au min
                'If unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran Then
                If (unYecran <= unMinYecran And unYecran >= unMaxYecran) Or (unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran) Then
                    uneZoneDessin.Line (unXecran, unYecran)-(unXecSuiv, unYecSuiv), uneCouleur
                End If
                'Stockage pour le segment suivant
                unXecran = unXecSuiv
                unYecran = unYecSuiv
            Next j
            
            'Calcul du dernier point pour la courbe temps/distance
            'ou la courbe vitesse/distance
            'Conversion des distances des décimètres au mètre
            unD2 = unParcours.maDistPar / 10 * unParcours.monCoefEta
            If unTypeCourbe = OngletCbeDT Then
                'Conversion du pas des dixièmes de secondes en minute
                unX2 = unParcours.maDuree / 600
            ElseIf unTypeCourbe = OngletCbeDV Then
                If unParcours.monLastPas = 0 Then
                    unX2 = 0
                Else
                    'Décimètre/dixième de seconde converti en km/h
                    unX2 = unParcours.monTabDist(unNbPoints) * unParcours.monCoefEta / unParcours.monLastPas * 3.6
                End If
            End If
            'Conversion en coordonnées écrans des coordonnées réelles
            'du point suivant
            unXecSuiv = ConvertirEnEcran(unMinXecran, unX2 - unMinXreel, uneDistMaxReelX, uneDistMaxEcranX)
            unYecSuiv = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD2, uneDistMaxReelY, uneDistMaxEcranY)
            'Dessin de la courbe segment par segment
            's'il est entre le min et le max y écran
            'min y écran  > max y écran car les y écran orientés vers le bas en Y,
            'donc aprés conversion donnée réelle en écran le max devient < au min
            'If unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran Then
            If (unYecran <= unMinYecran And unYecran >= unMaxYecran) Or (unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran) Then
                uneZoneDessin.Line (unXecran, unYecran)-(unXecSuiv, unYecSuiv), uneCouleur
            End If
        End If
    Next i
    
    'Dessin des lignes de rappel des repères en pointillé noir
    uneZoneDessin.DrawWidth = TraitFin
    uneZoneDessin.DrawStyle = vbDashDot
    If unPrint Then
        'On fixe le x de fin de dessin pour les traits de rappel des repères
        'en impression au x max plus la longueur du libellé en X et un petit décalage
        unXFinDessin = unMaxXecran + Printer.TextWidth(unLibX) + PicBoxTop
        unDecVertical = Printer.TextHeight("W")
    Else
        'On fixe le x de fin de dessin pour les traits de rappel des repères
        'sur l'écran au x de fin de la picture box
        unXFinDessin = unMaxXecran + uneMargeD
    End If
    
    unNbRep = uneForm.maColRepere.Count
    For i = 1 To unNbRep
        'Dessin de la ligne de rappel et du nom du repère
        's'il est entre le min et le max y écran
        'min y écran  > max y écran car les y écran orientés vers le bas en Y,
        'donc aprés conversion donnée réelle en écran le max devient < au min
        Set unRep = uneForm.maColRepere(i)
        unYecran = ConvertirEnEcran(unMaxYecran, unMaxYreel - unRep.monAbsCurv, uneDistMaxReelY, uneDistMaxEcranY)
        If unYecran <= unMinYecran And unYecran >= unMaxYecran Then
            uneZoneDessin.Line (unMinXecran, unYecran)-(unXFinDessin, unYecran), QBColor(0)
            If unPrint Then
                'Impression du nom court du repère
                Printer.CurrentX = unXFinDessin
                If i = unNbRep And unNbRep > 1 Then
                    'On évite le chevauchement des noms courts
                    'deux derniers repères en impressions
                    unYprec = ConvertirEnEcran(unMaxYecran, unMaxYreel - uneForm.maColRepere(unNbRep - 1).monAbsCurv, uneDistMaxReelY, uneDistMaxEcranY)
                    If unYprec - unYecran < unDecVertical Then
                        unDecHicon = unDecVertical
                        unDecVertical = unDecVertical * 2
                    End If
                End If
                Printer.CurrentY = unYecran - unDecVertical / 2
                uneZoneDessin.Print unRep.monNomCourt
                'unXpos = unXFinDessin + unMaxWidthNomRep 'Largeur maxi de noms courts
                unXpos = unXFinDessin + Printer.TextWidth(unRep.monNomCourt) + 0.1 * UnCmEnTwips '0.1 cm = 1 mm
                unYpos = unYecran - unDecHicon / 2 - unRep.monIcone.Height / 4 '2
                Printer.PaintPicture unRep.monIcone.Picture, unXpos, unYpos, unRep.monIcone.Width / 2, unRep.monIcone.Height / 2
            End If
        End If
    Next i
    
    'Dessin du repère XY avec graduations principales et secondaires
    uneZoneDessin.DrawWidth = TraitFin
    uneZoneDessin.DrawStyle = vbSolid
    uneZoneDessin.Line (unMinXecran, unMinYecran)-(unMaxXecran, unMinYecran), 0
    uneZoneDessin.Line (unMinXecran, unMinYecran)-(unMinXecran, unMaxYecran), 0
    
    'Affichage du libellé sur l'axe des x
    uneZoneDessin.CurrentX = unMaxXecran
    uneZoneDessin.CurrentY = unMinYecran - uneZoneDessin.TextHeight(unLibX)
    uneZoneDessin.Print unLibX
        
    'Restauration du pointeur souris par défaut
    uneForm.MousePointer = vbDefault
End Sub

Public Sub RemplirTabSS(uneForm As Form)
    'Remplissage du tableau de synthése et de statistiques
    Dim unNbRep As Integer, unRep As Repere, unRepNext As Repere
    Dim unNbParUtil As Integer, unPar As Parcours, unItiComplet As String
    Dim uneColRep As Collection, uneColVal As Collection, uneColRes As Collection
    Dim uneD1 As Long, uneD2 As Long, unIndFirstRep As Integer
    Dim unTPmin As Single, unTPmax As Single, unTPmoy As Single, unTP As Long
    Dim uneVmin As Single, uneVmax As Single, uneVmoy As Single, uneV As Single
    Dim unTAmin As Single, unTAmax As Single, unTAmoy As Single, unTA As Long
    Dim unNAmin As Single, unNAmax As Single, unNAmoy As Single, unNA As Long
    Dim unPAmin As Single, unPAmax As Single, unPAmoy As Single, unPA As Long
    Dim unTTmin As Single, unTTmax As Single, unTTmoy As Single, unTT As Long
    Dim unNTmin As Single, unNTmax As Single, unNTmoy As Single, unNT As Long
    Dim unPTmin As Single, unPTmax As Single, unPTmoy As Single, unPT As Long
    Dim unEcartType As Single, uneErreurAbs As Single
    
    'Pour effacer l'affichage précédent
    uneForm.SpreadTabSS.MaxRows = 0
    
    'Création d'une collection pour stocker les  valeurs par parcours
    Set uneColVal = New Collection
    'Création d'une collection pour stocker les  résultats min, max et moyen
    'par parcours
    Set uneColRes = New Collection
    
    'Création d'une autre collection de repères
    'mais triées par abscisse curviligne croissant
    unNbRep = uneForm.maColRepere.Count
    Set uneColRep = New Collection
    'Détermination des abscisses début et fin de la section de travail
    'de telle façon que Abs début < abs fin (donc D1 < D2)
    'Et stockage dans unIndFirstRep de l'indice du repère stocké en premier
    'dans la collection uneColRep qui contiendra les repères triés par ordre
    'croissant et compris entre abs début et fin
    If uneForm.CheckSection.Value = 1 Then
        'Cas où une section de travail est définie
        'Détermination des abscisses début et fin de la section
        unIndFirstRep = uneForm.ComboRepDebSec.ListIndex + 1
        Set unRep = uneForm.maColRepere(unIndFirstRep)
        Set unRepNext = uneForm.maColRepere(uneForm.ComboRepFinSec.ListIndex + 1)
        uneColRep.Add unRep
        If unRep.monAbsCurv < unRepNext.monAbsCurv Then
            uneD1 = unRep.monAbsCurv
            uneD2 = unRepNext.monAbsCurv
        Else
            uneD2 = unRep.monAbsCurv
            uneD1 = unRepNext.monAbsCurv
        End If
    Else
        'Section de travail = tout l'tinéraire
        'D'où des abscisses début et fin englobant tout largement
        uneColRep.Add uneForm.maColRepere(1)
        unIndFirstRep = 1
        uneD1 = -100
        uneD2 = 10000000
    End If
    
    'Tri croissant par insertion au fur et à mesure dans la collection uneColRep
    For i = 1 To unNbRep
        Set unRep = uneForm.maColRepere(i)
        If unRep.monAbsCurv >= uneD1 And unRep.monAbsCurv <= uneD2 And i <> unIndFirstRep Then
            'Cas où le repère en cours est entre les abscisses début et fin
            'de la section de travail ou de tout l'itinéraire
            'Et que ce n'est pas le premier repère mis dans uneColRep au départ
            For j = 1 To uneColRep.Count
                If unRep.monAbsCurv < uneColRep(j).monAbsCurv Then
                    unePos = j
                    'Mise en j ème position
                    uneColRep.Add unRep, , j
                    Exit For 'fin de boucle for j
                End If
            Next j
            If j = uneColRep.Count + 1 Then
                'Cas où plus grand que tous les éléments de la collection uneColRep
                '==> Mis en dernier position, car c'est le plus grand en cours
                'En VB, le compteur en sortie de for vaut Fin-du-Compteur + 1
                uneColRep.Add unRep
            End If
        End If
    Next i
        
    'Calcul du nombre de repères à prendre en compte
    unNbRep = uneColRep.Count
    'Calcul du nombre de parcours utilisés
    unNbParUtil = DonnerNbParcoursUtil(uneForm)
    
    If uneForm.CheckSection.Value = 1 And unNbRep = 2 Then
        'Cas où on est en section de travail et qu'il n'y a que deux repères
        'sur toutes la section, on n'affichera pas deux fois les infos entre
        'R1 et R2
        unNbRep = 1
        unSeulTroncon = True
        'Ainsi on ne fera que le cas i = 0
    Else
        unSeulTroncon = False
    End If
    
    '******************************************************
    'Calcul des valeurs mini, maxi et somme des valeurs
    'et stockage dans la collection uneColVal
    '******************************************************
    
    uneForm.SpreadTabSS.MaxRows = unNbRep * 9
    'multiple de 9 (= 8 +1) car huit valeurs à afficher par tronçon (cf colonne 2)
    'et le tronçon du parcours total plus une ligne blanche de séparation
    
    unItiComplet = ""
    For i = unNbRep - 1 To 0 Step -1
        If i = 0 Then
            'Cas du parcours complet, donc sur tout l'itinéraire ou la section
            Set unRep = uneColRep(1)
            If unSeulTroncon Then
                Set unRepNext = uneColRep(2)
            Else
                Set unRepNext = uneColRep(unNbRep)
                unItiComplet = "Itinéraire complet"
            End If
            'On fixe les bornes de travail en distance en décimètres
            If uneForm.CheckSection.Value = 1 Then
                'Cas où une section de travail est définie
                uneD1 = unRep.monAbsCurv * 10
                uneD2 = unRepNext.monAbsCurv * 10
            Else
                uneD1 = -1000
                uneD2 = 1000000
            End If
        Else
            'Cas d'un tronçon
            Set unRep = uneColRep(i)
            Set unRepNext = uneColRep(i + 1)
            'On fixe les bornes de travail en distance en décimètres
            uneD1 = unRep.monAbsCurv * 10
            uneD2 = unRepNext.monAbsCurv * 10
        End If
        
        'Initialisation
        unTPmin = 10000000
        unTPmax = 0
        unTPmoy = 0
        
        uneVmin = 10000000
        uneVmax = 0
        uneVmoy = 0
        
        unTAmin = 10000000
        unTAmax = 0
        unTAmoy = 0
        
        unNAmin = 10000000
        unNAmax = 0
        unNAmoy = 0
        
        unPAmin = 10000000
        unPAmax = 0
        unPAmoy = 0
        
        unTTmin = 10000000
        unTTmax = 0
        unTTmoy = 0
        
        unNTmin = 1000000
        unNTmax = 0
        unNTmoy = 0
        
        unPTmin = 10000000
        unPTmax = 0
        unPTmoy = 0
        
        unNbParTotal = uneForm.maColParcours.Count
        j = 0
        For k = 1 To unNbParTotal
            Set unPar = uneForm.maColParcours(k)
            If unPar.monIsUtil Then
                j = j + 1
                'Calcul des infos du tronçon pour le parcours j
                'Les distances sont en décimètres pour les calculs d'où le * 10
                unPar.CalculerLesVitDistDureeEtArrets uneD1, uneD2
                unPar.CalculerNbEtDureeDoubleTop uneD1, uneD2
                
                'Stockage des valeurs du parcours sur le tronçon
                unTP = unPar.monTFinSection - unPar.monTDebSection
                If unTPmin > unTP Then unTPmin = unTP
                If unTPmax < unTP Then unTPmax = unTP
                unTPmoy = (unTPmoy * (j - 1) + unTP) / j
                
                uneV = unPar.maVmoy
                If uneVmin > unPar.maVmoy Then uneVmin = unPar.maVmoy
                If uneVmax < unPar.maVmoy Then uneVmax = unPar.maVmoy
                uneVmoy = (uneVmoy * (j - 1) + unPar.maVmoy) / j
                
                unTA = unPar.monTpsArret
                If unTAmin > unPar.monTpsArret Then unTAmin = unPar.monTpsArret
                If unTAmax < unPar.monTpsArret Then unTAmax = unPar.monTpsArret
                unTAmoy = (unTAmoy * (j - 1) + unPar.monTpsArret) / j
                
                unNA = unPar.monNbArret
                If unNAmin > unPar.monNbArret Then unNAmin = unPar.monNbArret
                If unNAmax < unPar.monNbArret Then unNAmax = unPar.monNbArret
                unNAmoy = (unNAmoy * (j - 1) + unPar.monNbArret) / j
                
                If unPar.monTFinSection = unPar.monTDebSection Then
                    unPA = 0
                Else
                    unPA = unPar.monTpsArret / (unPar.monTFinSection - unPar.monTDebSection) * 100
                End If
                If unPAmin > unPA Then unPAmin = unPA
                If unPAmax < unPA Then unPAmax = unPA
                unPAmoy = (unPAmoy * (j - 1) + unPA) / j
                
                unTT = unPar.monTpsDbTop
                If unTTmin > unPar.monTpsDbTop Then unTTmin = unPar.monTpsDbTop
                If unTTmax < unPar.monTpsDbTop Then unTTmax = unPar.monTpsDbTop
                unTTmoy = (unTTmoy * (j - 1) + unPar.monTpsDbTop) / j
                
                unNT = unPar.monNbDbTop
                If unNTmin > unPar.monNbDbTop Then unNTmin = unPar.monNbDbTop
                If unNTmax < unPar.monNbDbTop Then unNTmax = unPar.monNbDbTop
                unNTmoy = (unNTmoy * (j - 1) + unPar.monNbDbTop) / j
                
                If unPar.monTFinSection = unPar.monTDebSection Then
                    unPT = 0
                Else
                    unPT = unPar.monTpsDbTop / (unPar.monTFinSection - unPar.monTDebSection) * 100
                End If
                If unPTmin > unPT Then unPTmin = unPT
                If unPTmax < unPT Then unPTmax = unPT
                unPTmoy = (unPTmoy * (j - 1) + unPT) / j
                'Stockage dans la collection des valeurs en première position
                If j = 1 Then
                    'Cas où uneColVal est vide, sinon le add avec ,,1 plante
                    uneColVal.Add unPT
                Else
                    uneColVal.Add unPT, , 1
                End If
                uneColVal.Add unNT, , 1
                uneColVal.Add unTT, , 1
                uneColVal.Add unPA, , 1
                uneColVal.Add unNA, , 1
                uneColVal.Add unTA, , 1
                uneColVal.Add uneV, , 1
                uneColVal.Add unTP, , 1
            End If
        Next k
        
        'Stockage dans la collection des résultats mini, maxi et moyen
        'en première position avec formattage en texte avec le libellé
        'de la valeur
        uneColRes.Add unPTmoy 'Au début unecolres vide ==> add ,,1 plante
        uneColRes.Add unPTmax, , 1
        uneColRes.Add unPTmin, , 1
        uneColRes.Add "% Temps dble Top", , 1
        
        uneColRes.Add unNTmoy, , 1
        uneColRes.Add unNTmax, , 1
        uneColRes.Add unNTmin, , 1
        uneColRes.Add "Nbre double Top", , 1
        
        uneColRes.Add unTTmoy, , 1
        uneColRes.Add unTTmax, , 1
        uneColRes.Add unTTmin, , 1
        uneColRes.Add "Temps double Top", , 1
        
        uneColRes.Add unPAmoy, , 1
        uneColRes.Add unPAmax, , 1
        uneColRes.Add unPAmin, , 1
        uneColRes.Add "% Temps d'arrêts", , 1
        
        uneColRes.Add unNAmoy, , 1
        uneColRes.Add unNAmax, , 1
        uneColRes.Add unNAmin, , 1
        uneColRes.Add "Nombre d'arrêts", , 1
        
        uneColRes.Add unTAmoy, , 1
        uneColRes.Add unTAmax, , 1
        uneColRes.Add unTAmin, , 1
        uneColRes.Add "Temps d'arrêts", , 1
        
        uneColRes.Add uneVmoy, , 1
        uneColRes.Add uneVmax, , 1
        uneColRes.Add uneVmin, , 1
        uneColRes.Add "V moyenne (km/h)", , 1
        
        uneColRes.Add unTPmoy, , 1
        uneColRes.Add unTPmax, , 1
        uneColRes.Add unTPmin, , 1
        uneColRes.Add "Temps parcours", , 1
        
        'Calcul de l'écart type de chaque valeur
        'et de l'erreur absolue, la 1er moyenne en 4 ème place dans uneColRes
        'puis avec l'ajout dans uneColRes de l'écart type et de l'erreur absolue
        'en place multiple de 6
        j0 = 0
        j = 4
        unNbRes = uneColRes.Count
        Do
            uneValMoy = uneColRes(j)
            unEcartType = 0
            uneErreurAbs = 0
            j0 = j0 + 1
            For k = j0 To uneColVal.Count Step 8
                uneVal = uneColVal(k)
                'Somme des carrés des valeurs
                unEcartType = unEcartType + uneVal * uneVal
                'Somme des écarts absolus entre la valeur et la moyenne
                uneErreurAbs = uneErreurAbs + Abs(uneVal - uneValMoy)
            Next k
            unEcartType = Sqr(Abs(unEcartType / unNbParUtil - uneValMoy * uneValMoy))
            uneErreurAbs = uneErreurAbs / unNbParUtil
            'Insertion dans la collection des résultats
            'après les valeurs min, max et moyenne, la moyenne est en 3ème place
            If j = unNbRes Then
                'Cas de l'insertion du dernier ecart type et erreur absolue
                'on le met à la fin
                 uneColRes.Add unEcartType
                uneColRes.Add uneErreurAbs
            Else
                'Autres cas : on les met après les valeurs moyennes
                uneColRes.Add unEcartType, , j + 1
                uneColRes.Add uneErreurAbs, , j + 2
            End If
            'Incrémentation pour les coups suivants
            j = j + 6
            unNbRes = unNbRes + 2
        Loop Until j > unNbRes
        
        'On vide la collection stockant les valeurs du tronçon
        ViderCollection uneColVal
        
        'Affichage dans le tableau de synthèse et stat
        j0 = 0
        For j = 9 * i + 1 To 9 * (i + 1)
            If j = 1 Then
                j0 = 1
                If unItiComplet <> "" Then
                    uneForm.SpreadTabSS.Row = 1
                    uneForm.SpreadTabSS.Col = 1
                    uneForm.SpreadTabSS.Text = unItiComplet
                End If
            End If
            
            'Affichage des info sur le tronçon le nom court puis les abs
            'du repère, quand les noms courts étaient sur 10 caractères
            'If j Mod 9 = 1 Then
            '    uneForm.SpreadTabSS.Row = j + j0
            '    uneForm.SpreadTabSS.Col = 1
            '    uneForm.SpreadTabSS.Text = "De " + unRep.monNomCourt
            'ElseIf j Mod 9 = 2 Then
            '    uneForm.SpreadTabSS.Row = j + j0
            '    uneForm.SpreadTabSS.Col = 1
            '    uneForm.SpreadTabSS.Text = "Abs = " + Format(unRep.monAbsCurv) + " m"
            'ElseIf j Mod 9 = 3 Then
            '    uneForm.SpreadTabSS.Row = j + j0
            '    uneForm.SpreadTabSS.Col = 1
            '    uneForm.SpreadTabSS.Text = "à " + unRepNext.monNomCourt
            'ElseIf j Mod 9 = 4 Then
            '    uneForm.SpreadTabSS.Row = j + j0
            '    uneForm.SpreadTabSS.Col = 1
            '    uneForm.SpreadTabSS.Text = "Abs = " + Format(unRepNext.monAbsCurv) + " m"
                
            'Affichage des info sur le tronçon Abs puis le nom court
            'du repère, changement du au passage des noms courts de 10
            'à 15 caractères
            If j Mod 9 = 1 Then
                uneForm.SpreadTabSS.Row = j + j0
                uneForm.SpreadTabSS.Col = 1
                uneForm.SpreadTabSS.Text = "De l'abs = " + Format(unRep.monAbsCurv) + " m"
            ElseIf j Mod 9 = 2 Then
                uneForm.SpreadTabSS.Row = j + j0
                uneForm.SpreadTabSS.Col = 1
                uneForm.SpreadTabSS.Text = unRep.monNomCourt
            ElseIf j Mod 9 = 3 Then
                uneForm.SpreadTabSS.Row = j + j0
                uneForm.SpreadTabSS.Col = 1
                uneForm.SpreadTabSS.Text = "à l'abs = " + Format(unRepNext.monAbsCurv) + " m"
            ElseIf j Mod 9 = 4 Then
                uneForm.SpreadTabSS.Row = j + j0
                uneForm.SpreadTabSS.Col = 1
                uneForm.SpreadTabSS.Text = unRepNext.monNomCourt
            ElseIf j Mod 9 = 5 Then
                uneForm.SpreadTabSS.Row = j + j0
                uneForm.SpreadTabSS.Col = 1
                uneForm.SpreadTabSS.Text = "Long = " + Format(unRepNext.monAbsCurv - unRep.monAbsCurv) + " m"
            End If
            
            unePos = j Mod 9 - 1
            'Affichage de la colonne 2
            'Dans uneColRes 6 valeurs pour chacune des 8 infos à afficher
            uneForm.SpreadTabSS.Row = j
            uneForm.SpreadTabSS.Col = 2
            If unePos = -1 Then
                uneForm.SpreadTabSS.Text = ""
            Else
                uneForm.SpreadTabSS.Text = uneColRes(unePos * 6 + 1)
            End If
            For k = 3 To uneForm.SpreadTabSS.MaxCols
                'Autres colonnes
                uneForm.SpreadTabSS.Col = k
                'Dans uneColRes 6 valeurs pour chacune des 8 infos à afficher
                If unePos > -1 Then uneVal = uneColRes(unePos * 6 + k - 1)
                Select Case unePos + 1
                    Case 0
                        'Ligne vide de séparation
                        uneForm.SpreadTabSS.Text = ""
                    Case 1, 3, 6
                        'Temps converti et arrondi à la seconde
                        'et formatter en XXh YYmn ZZs
                        uneForm.SpreadTabSS.Text = FormatterTempsEnHMNS(CLng(uneVal))
                    Case 2
                        'Single (vitesse) formatter en XXX.YY
                        uneForm.SpreadTabSS.Text = Format(uneVal, "fixed")
                    Case 5, 8
                        '% Temps formatter en XXX%
                        uneForm.SpreadTabSS.Text = Format(CLng(uneVal)) + "%"
                    Case 4, 7
                        'Entier (Nombre d'arrêts et de double top) formatter en texte
                        If uneVal = Int(uneVal) Then
                            uneForm.SpreadTabSS.Text = Format(uneVal)
                        Else
                            uneForm.SpreadTabSS.Text = Format(uneVal, "fixed")
                        End If
                    Case Else
                        MsgBox MsgErreurProg + "Numéro de colonne inconnue dans RemplirTabSS", vbCritical
                End Select
            Next k
        Next j
        'On vide la collection des valeurs résultats formattées
        ViderCollection uneColRes
    Next i
    
    ViderCollection uneColRep
    Set uneColRep = Nothing
    Set uneColRes = Nothing
    Set uneColVal = Nothing
End Sub

Public Sub RemplirTabBrut(uneForm As Form)
    'Remplissage du tableau brut
    Dim unNbRep As Integer, unRep As Repere, unRepNext As Repere
    Dim unNbParUtil As Integer, unPar As Parcours, uneColRep As Collection
    Dim uneD1 As Long, uneD2 As Long
    Dim unIndFirstRep As Integer
    
    'Pour effacer l'affichage précédent
    uneForm.SpreadTabBrut.MaxRows = 0
    
    'Création d'une autre collection de repères
    'mais triées par abscisse curviligne croissant
    unNbRep = uneForm.maColRepere.Count
    Set uneColRep = New Collection
    'Détermination des abscisses début et fin de la section de travail
    'de telle façon que Abs début < abs fin (donc D1 < D2)
    'Et stockage dans unIndFirstRep de l'indice du repère stocké en premier
    'dans la collection uneColRep qui contiendra les repères triés par ordre
    'croissant et compris entre abs début et fin
    If uneForm.CheckSection.Value = 1 Then
        'Cas où une section de travail est définie
        'Détermination des abscisses début et fin de la section
        unIndFirstRep = uneForm.ComboRepDebSec.ListIndex + 1
        Set unRep = uneForm.maColRepere(unIndFirstRep)
        Set unRepNext = uneForm.maColRepere(uneForm.ComboRepFinSec.ListIndex + 1)
        uneColRep.Add unRep
        If unRep.monAbsCurv < unRepNext.monAbsCurv Then
            uneD1 = unRep.monAbsCurv
            uneD2 = unRepNext.monAbsCurv
        Else
            uneD2 = unRep.monAbsCurv
            uneD1 = unRepNext.monAbsCurv
        End If
    Else
        'Section de travail = tout l'tinéraire
        'D'où des abscisses début et fin englobant tout largement
        uneColRep.Add uneForm.maColRepere(1)
        unIndFirstRep = 1
        uneD1 = -100
        uneD2 = 10000000
    End If
    
    'Tri croissant par insertion au fur et à mesure dans la collection uneColRep
    For i = 1 To unNbRep
        Set unRep = uneForm.maColRepere(i)
        If unRep.monAbsCurv >= uneD1 And unRep.monAbsCurv <= uneD2 And i <> unIndFirstRep Then
            'Cas où le repère en cours est entre les abscisses début et fin
            'de la section de travail ou de tout l'itinéraire
            'Et que ce n'est pas le premier repère mis dans uneColRep au départ
            For j = 1 To uneColRep.Count
                If unRep.monAbsCurv < uneColRep(j).monAbsCurv Then
                    unePos = j
                    'Mise en j ème position
                    uneColRep.Add unRep, , j
                    Exit For 'fin de boucle for j
                End If
            Next j
            If j = uneColRep.Count + 1 Then
                'Cas où plus grand que tous les éléments de la collection uneColRep
                '==> Mis en dernier position, car c'est le plus grand en cours
                'En VB, le compteur en sortie de for vaut Fin-du-Compteur + 1
                uneColRep.Add unRep
            End If
        End If
    Next i
        
    'Calcul du nombre de repères à prendre en compte
    unNbRep = uneColRep.Count
    'Calcul du nombre de parcours utilisés
    unNbParUtil = DonnerNbParcoursUtil(uneForm)
    
    If uneForm.CheckSection.Value = 1 And unNbRep = 2 Then
        'Cas où on est en section de travail et qu'il n'y a que deux repères
        'sur toutes la section, on n'affichera pas deux fois les infos entre
        'R1 et R2
        unNbRep = 1
        unSeulTroncon = True
        'Ainsi on ne fera que le cas i = 0
    Else
        unSeulTroncon = False
    End If
    'Calcul et remplissage du tableau brut,
    'une ligne par parcours pour chaque tronçon plus l'itinéraire complet
    uneForm.SpreadTabBrut.MaxRows = unNbRep * unNbParUtil
    
    '1ère ligne les info sur l'itinéraire complet
    'Les autres lignes, les infos sur chaque tronçons entre repères consécutifs
    For i = unNbRep - 1 To 0 Step -1
        If i = 0 Then
            'Cas du parcours complet, donc sur tout l'itinéraire ou la section
            Set unRep = uneColRep(1)
            If unSeulTroncon Then
                Set unRepNext = uneColRep(2)
                unItiComplet = ""
            Else
                Set unRepNext = uneColRep(unNbRep)
                unItiComplet = "Itinéraire complet" + Chr(13)
            End If
            'On fixe les bornes de travail en distance en décimètres
            If uneForm.CheckSection.Value = 1 Then
                'Cas où une section de travail est définie
                uneD1 = unRep.monAbsCurv * 10
                uneD2 = unRepNext.monAbsCurv * 10
            Else
                uneD1 = -1000
                uneD2 = 1000000
            End If
        Else
            'Cas d'un tronçon
            Set unRep = uneColRep(i)
            Set unRepNext = uneColRep(i + 1)
            unItiComplet = ""
            'On fixe les bornes de travail en distance en décimètres
            uneD1 = unRep.monAbsCurv * 10
            uneD2 = unRepNext.monAbsCurv * 10
        End If
        
        unNbParTotal = uneForm.maColParcours.Count
        j = 0
        uneNewLigneTronçon = True
        For k = 1 To unNbParTotal
            Set unPar = uneForm.maColParcours(k)
            If unPar.monIsUtil Then
                j = j + 1
                uneString = ""
                'Première colonne rien ou info tronçon
                uneForm.SpreadTabBrut.Col = 1
                uneForm.SpreadTabBrut.Row = i * unNbParUtil + j
                If uneNewLigneTronçon Then
                    uneNewLigneTronçon = False
                    'Affichage des info sur le tronçon quand les noms courts
                    'de repères étaient de 10 caractères(nom court rep puis abs)
                    'uneString = unItiComplet + "De " + unRep.monNomCourt
                    'uneString = uneString + Chr(13) + "Abs = " + Format(unRep.monAbsCurv, "# ### ##0") + " m"
                    'uneString = uneString + Chr(13) + "à " + unRepNext.monNomCourt
                    'uneString = uneString + Chr(13) + "Abs = " + Format(unRepNext.monAbsCurv, "# ### ##0") + " m"
                    
                    'Affichage des info sur le tronçon Abs puis le nom court
                    'du repère, changement du au passage des noms courts de 10
                    'à 15 caractères
                    uneString = unItiComplet + "De l'abs = " + Format(unRep.monAbsCurv, "# ### ##0") + " m"
                    uneString = uneString + Chr(13) + unRep.monNomCourt
                    uneString = uneString + Chr(13) + "à l'abs = " + Format(unRepNext.monAbsCurv, "# ### ##0") + " m"
                    uneString = uneString + Chr(13) + unRepNext.monNomCourt
                End If
                uneForm.SpreadTabBrut.Text = uneString
                
                'Calcul des infos du tronçon pour le parcours j
                'Les distances sont en décimètres pour les calculs d'où le * 10
                unPar.CalculerLesVitDistDureeEtArrets uneD1, uneD2
                unPar.CalculerNbEtDureeDoubleTop uneD1, uneD2
                
                'Affichage du nom et date de la mesure du parcours en cours
                uneForm.SpreadTabBrut.Col = 2
                uneForm.SpreadTabBrut.Row = i * unNbParUtil + j
                uneString = unPar.monNom + Chr(13) + Mid(unPar.monJourSemaine, 1, 2)
                uneString = uneString + " " + Format(unPar.maDate)
                uneString = uneString + " " + Format(unPar.monHeureDebut)
                uneForm.SpreadTabBrut.Text = uneString
                
                'Affichage des autres infos du parcours sur le tronçon
                uneForm.SpreadTabBrut.Col = 3
                uneForm.SpreadTabBrut.Row = i * unNbParUtil + j
                uneString = Format(CLng(unPar.maDistParSection / 10), "# ### ##0") + " m"
                uneString = uneString + Chr(13) + FormatterTempsEnHMNS(unPar.monTFinSection - unPar.monTDebSection)
                uneForm.SpreadTabBrut.Text = uneString
                uneForm.SpreadTabBrut.Col = 4
                uneForm.SpreadTabBrut.Row = i * unNbParUtil + j
                uneForm.SpreadTabBrut.Text = Format(unPar.maVmoy, "fixed")
                uneForm.SpreadTabBrut.Col = 5
                uneForm.SpreadTabBrut.Row = i * unNbParUtil + j
                uneString = FormatterTempsEnHMNS(unPar.monTpsArret) + Chr(13)
                uneString = uneString + Format(unPar.monNbArret) + Chr(13)
                If unPar.monTFinSection = unPar.monTDebSection Then
                    uneString = uneString + "0%"
                Else
                    uneString = uneString + Format(CLng(unPar.monTpsArret / (unPar.monTFinSection - unPar.monTDebSection) * 100)) + "%"
                End If
                uneForm.SpreadTabBrut.Text = uneString
                uneForm.SpreadTabBrut.Col = 6
                uneForm.SpreadTabBrut.Row = i * unNbParUtil + j
                uneString = FormatterTempsEnHMNS(unPar.monTpsDbTop) + Chr(13)
                uneString = uneString + Format(unPar.monNbDbTop) + Chr(13)
                If unPar.monTFinSection = unPar.monTDebSection Then
                    uneString = uneString + "0%"
                Else
                    uneString = uneString + Format(CLng(unPar.monTpsDbTop / (unPar.monTFinSection - unPar.monTDebSection) * 100)) + "%"
                End If
                uneForm.SpreadTabBrut.Text = uneString
            End If
        Next k
    Next i
    
    ViderCollection uneColRep
    Set uneColRep = Nothing
End Sub

Public Sub DessinerSynoV(uneForm As Form, uneZoneDessin As Object)
    'Dessin du synotique des vitesses sur une zone
    'de dessin écran ou imprimante d'une étude (= uneForm)
    '(écran = picture box de l'onglet Synotique des vitesses)
    Dim uneMargeG As Single, uneMargeD As Single 'marges gauche et droite
    Dim uneMargeH As Single, uneMargeB As Single 'marges haut et bas
    Dim unLibX As String, uneHText As Single, uneWText As Single
    Dim uneCouleur As Long, uneV As Single
    Dim unMaxXecran As Single, unMaxYecran As Single
    Dim unMinXecran As Single, unMinYecran As Single
    Dim unMaxXreel As Single, unMinXreel As Single
    Dim unMaxYreel As Single, unMinYreel As Single
    Dim unYecran As Single, unXFinDessin As Single
    Dim unYecSuiv As Single, unYecPred As Single
    Dim unPasYGrad1 As Long, unPasYGrad2 As Long, unMaxWidthNomRep As Single
    Dim j As Long, unNbPoints As Long, unRep As Repere
    Dim uneDistMaxReelX As Single, uneDistMaxEcranX As Single
    Dim uneDistMaxReelY As Single, uneDistMaxEcranY As Single
    Dim unParcours As Parcours, unTabInfo(0 To 7) As String * 2
    
    'Affichage du sablier en pointeur souris pour symboliser l'attente
    uneForm.MousePointer = vbHourglass
    unPrint = (TypeOf uneZoneDessin Is Printer) '=-1 si impression, 0 sinon
    
    'Test si on affiche ou imprime plus de 17 parcours, car taille mini
    '= 8160 twips < largeur d'un écran 800x600 et largeur papier A4 portrait (valant environ 10000)
    'et pour chaque parcours on affichera des rectangles d'épaisseur = 240 twips
    'avec un espacement entre deux empilements de rectangles de parcours de la
    'même valeur = 240 twips
    'en effet 17 * 240 *2 = 8160
    unNbPar = 0
    For i = 1 To uneForm.maColParcours.Count
        If uneForm.maColParcours(i).monIsUtil Then unNbPar = unNbPar + 1
    Next i
    If unNbPar > 17 Then
        MsgBox "Impossible d'afficher ou d'imprimer plus de 17 parcours dans le synoptique des vitesses. Diminuer votre nombre de parcours utilisés.", vbExclamation
        uneForm.MousePointer = vbDefault
        Exit Sub
    End If
    If unNbPar > 8 Then
        'Si entre 9 et 17 parcours
        'unWRect = 240 ' = 8160 / 17 /2
        unWRect = 8160 / unNbPar / 2
        'Pour avoir une meilleure répartition des largeurs
    Else
        'Si entre 1 et 8 parcours
        unWRect = 510 ' = 8160 / 8 / 2
    End If
    
    'Initialisation de variables utilisées ci-après
    uneZoneDessin.Font.Name = "Arial"
    uneZoneDessin.Font.Size = 7
    uneZoneDessin.ForeColor = QBColor(0)
    uneHText = uneZoneDessin.TextHeight("9")
    uneLgChiffre = uneZoneDessin.TextWidth("9")
    uneWText = uneZoneDessin.TextWidth("W")
    unLibX = "Parcours"
    
    'On fixe les tailles de marges
    If TypeOf uneZoneDessin Is Printer Then
        unMaxWidthNomRep = FixerMargesImprimante(uneMargeG, uneMargeD, uneMargeH, uneMargeB)
    ElseIf TypeOf uneZoneDessin Is PictureBox Then
        uneZoneDessin.Cls
        FixerMargesPicBox uneForm, uneZoneDessin, uneMargeG, uneMargeD, uneMargeH, uneMargeB
    Else
        MsgBox MsgErreurProg + MsgErreurTypeZoneDessinInconnu + MsgIn + "ModuleMain:DessinerSynoV", vbCritical
        Exit Sub
    End If
    
    'Stockage des y min et max réels
    unMaxYreel = uneForm.monMaxD
    unMinYreel = uneForm.monMinD
    unMinXreel = uneMargeG
    unMaxXreel = uneZoneDessin.Width - uneMargeG
    
    'Calcul des pas de graduations primaires et secondaires et arrondis
    'de la valeur mini à la graduation secondaire juste inférieure
    'et de la valeur maxi à la graduation secondaire juste supérieure
    TrouverPasGradEtModifierMinMax unPasYGrad1, unPasYGrad2, unMaxYreel, unMinYreel
    
    'Stockage des y min et max réels en distances si on m'imprime pas
    If unPrint = False Then
        uneForm.monMaxD = unMaxYreel
        uneForm.monMinD = unMinYreel
        'Stockage des pas de graduations en distance
        uneForm.monPasGrad1 = unPasYGrad1
        uneForm.monPasGrad2 = unPasYGrad2
    End If
                
    'Variables servant pour la conversion coordonnées réelles en écran
    uneDistMaxReelX = unMaxXreel - unMinXreel
    uneDistMaxEcranX = uneZoneDessin.Width - uneMargeG - uneMargeD
    uneDistMaxReelY = unMaxYreel - unMinYreel
    uneDistMaxEcranY = uneZoneDessin.Height - uneMargeH - uneMargeB

    'Conversion en coordonnées Y écran des distances réelles
    unMaxYecran = uneMargeH
    unMinYecran = ConvertirEnEcran(unMaxYecran, unMaxYreel - unMinYreel, unMaxYreel - unMinYreel, uneZoneDessin.Height - uneMargeH - uneMargeB)

    'Conversion en coordonnées X écran des temps ou vitesses réels
    unMinXecran = uneMargeG
    unMaxXecran = ConvertirEnEcran(unMinXecran, unMaxXreel - unMinXreel, unMaxXreel - unMinXreel, uneZoneDessin.Width - uneMargeG - uneMargeD)
    
    'Dessin en trait continu fin
    uneZoneDessin.DrawWidth = TraitFin
    uneZoneDessin.DrawStyle = vbSolid
    
    'Mise des graduations sur OY
    unAfficheToutY = ((unMaxYreel - unMinYreel) < 26 * unPasYGrad2)
    For i = unMinYreel To unMaxYreel Step unPasYGrad2
        'Dessin du trait de graduation primaire ou secondaire
        unYecran = ConvertirEnEcran(unMaxYecran, unMaxYreel - i, unMaxYreel - unMinYreel, uneZoneDessin.Height - uneMargeH - uneMargeB)
        uneZoneDessin.Line (unMinXecran, unYecran)-(unMinXecran - 60 * (1 - (i Mod unPasYGrad1 = 0)), unYecran), 0
        
        'Placement en X écran
        uneZoneDessin.CurrentX = unMinXecran - uneZoneDessin.TextWidth(Format(i)) - 120
        
        If i Mod unPasYGrad1 = 0 Or unAfficheToutY Then
            'Affichage de la valeur de la graduation primaire
            'ou Affichage des valeurs des graduations secondaires si au plus
            '26 sont à afficher, contrainte de 26 due aux écrans 800x600
            uneZoneDessin.CurrentY = unYecran - uneZoneDessin.TextHeight(Format(i)) / 2
            uneZoneDessin.Print Format(i)
        ElseIf i = unMaxYreel Then
            'Affichage de la valeur max
            'Si le max est une graduation primaire ou secondaire,
            'sa position d'affichage est gérée par le 1er if ci-dessus.
            'Sinon l'affichage du max est décalée vers le haut
            'pour éviter les superpositions avec les graduations
            'Calcul du Y écran de la graduation précédente
            unYecPred = ConvertirEnEcran(unMaxYecran, unMaxYreel - i + unPasYGrad2, unMaxYreel - unMinYreel, uneZoneDessin.Height - uneMargeH - uneMargeB)
            'Affichage de la valeur max avec ajout d'un décalage si impression
            unDecalage = Abs(unYecPred - unYecran < uneHText) '=1 si vrai, 0 si faux
            uneZoneDessin.CurrentY = (unYecran - uneHText / 2) * (1 - unDecalage) + (unYecPred - uneHText * (1.2 - unPrint * 0.2)) * unDecalage
            uneZoneDessin.Print Format(i)
        ElseIf i = unMinYreel Then
            'Affichage de la valeur min
            'Si le min est une graduation primaire ou secondaire,
            'sa position d'affichage est gérée par le 1er if ci-dessus.
            'Sinon l'affichage du min est décalée vers le bas.
            'pour éviter les superpositions avec les graduations
            'Calcul du Y écran de la graduation suivante
            unYecSuiv = ConvertirEnEcran(unMaxYecran, unMaxYreel - i - unPasYGrad2, unMaxYreel - unMinYreel, uneZoneDessin.Height - uneMargeH - uneMargeB)
            'Affichage de la valeur min avec ajout d'un décalage si impression
            unDecalage = Abs(unYecran - unYecSuiv < uneHText) '=1 si vrai, 0 si faux
            uneZoneDessin.CurrentY = (unYecran - uneHText / 2) * (1 - unDecalage) + (unYecSuiv + uneHText * (0.2 - unPrint * 0.2)) * unDecalage
            uneZoneDessin.Print Format(i)
        End If
    Next i
    
    'Affichage de la valeur et son unité affichée sur Oy
    If unPrint Then
        uneZoneDessin.CurrentX = unMinXecran - uneZoneDessin.TextWidth("Distance (m)") / 2
        uneZoneDessin.CurrentY = unMaxYecran - uneHText * 2
        uneZoneDessin.Print "Distance (m)"
    End If
        
    'Dessin des rectangles de couleur des classes de vitesse de chacun
    'des parcours emplilés verticalement pour chaque parcours avec Pnum-parcours
    'sous l'axe Ox de la couleur du parcours pour repèrer chacun des parcours
    unDecW = unWRect
    unX1 = uneMargeG - unDecW
    For i = 1 To uneForm.maColParcours.Count
        Set unParcours = uneForm.maColParcours(i)
        If unParcours.monIsUtil Then
            uneCouleur = unParcours.maCouleur
            unNbPoints = unParcours.monNbPas
            
            'Positionnement en X des numéros et des synoptiques de vitesses
            unX1 = unX1 + unDecW * 2
            unX2 = unX1 + unDecW
            
            'Affichage du libelle identifiant le parcours sous Ox
            uneZoneDessin.ForeColor = uneCouleur
            unIdPar = "P" + Format(i - 1) 'car 0 = parcours moyen = 1er parcours
            uneZoneDessin.CurrentX = unX1 + (unWRect - uneZoneDessin.TextWidth(unIdPar)) / 2
            uneZoneDessin.CurrentY = unMinYecran + 60
            uneZoneDessin.Print unIdPar
            
            'Affichage de jour, date et heure verticalement 2 caractères par lignes
            uneZoneDessin.ForeColor = QBColor(0)
            unTabInfo(0) = " "
            unTabInfo(1) = Mid(unParcours.monJourSemaine, 1, 1) + LCase(Mid(unParcours.monJourSemaine, 2, 1))
            unTabInfo(2) = Mid(unParcours.maDate, 1, 2)
            unTabInfo(3) = Mid(unParcours.maDate, 4, 2)
            unTabInfo(4) = Mid(unParcours.maDate, 9, 2)
            unTabInfo(5) = "à"
            unTabInfo(6) = Mid(unParcours.monHeureDebut, 1, 2)
            unTabInfo(7) = Mid(unParcours.monHeureDebut, 4, 2)
            For j = 0 To 7
                uneZoneDessin.CurrentX = unX1 - uneLgChiffre * 2 - 30
                uneZoneDessin.CurrentY = unMinYecran - uneHText * (j + 1)
                uneZoneDessin.Print unTabInfo(7 - j)
            Next j
            'Affichage du nom du parcours verticalement à gauche
            'de son synoptique de vitesses
            uneZoneDessin.ForeColor = QBColor(0)
            uneLgNom = Len(unParcours.monNom)
            For j = uneLgNom To 1 Step -1
                uneZoneDessin.CurrentX = unX1 - uneWText - 30
                uneZoneDessin.CurrentY = unMinYecran - (uneLgNom - j + 9) * uneHText
                '7 car il y a 7 lignes d'info avant
                uneZoneDessin.Print Mid(unParcours.monNom, j, 1)
            Next j
            
            '*******************************************************
            'Affichage des rectangles de classes de vitesses
            '*******************************************************
        
            'Récup des données pour calculer les vitesses
            unD1 = 0
            
            'Conversion des distances des décimètres au mètre
            unD2 = unParcours.monTabDist(1) / 10 * unParcours.monCoefEta
            If unParcours.monFirstPas = 0 Then
                uneV = 0
            Else
                'mètre/dixième de seconde converti en km/h
                uneV = unD2 / unParcours.monFirstPas * 36
            End If
            'Récupération de la couleur suivant l'appartenance à
            'telle ou telle classe de vitesses
            uneCouleur = DonnerCouleurClasseV(uneV)
            
            'Conversion en coordonnées écrans des distances réelles (les D)
            'du premier point =(unX1, unD1) et du deuxième point =(unX2, unD2)
            'Les X1 et X2 sont des coordonnées écran ==> on ne le convertit pas
            unYecran = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD1, uneDistMaxReelY, uneDistMaxEcranY)
            unYecSuiv = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD2, uneDistMaxReelY, uneDistMaxEcranY)
            'Dessin de la courbe du premier segment
            's'il est entre le min et le max y écran
            'min y écran  > max y écran car les y écran orientés vers le bas en Y,
            'donc aprés conversion donnée réelle en écran le max devient < au min
            'If unYecran <= unMinYecran And unYecran >= unMaxYecran Then
            If (unYecran <= unMinYecran And unYecran >= unMaxYecran) Or (unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran) Then
                uneZoneDessin.Line (unX1, unYecran)-(unX2, unYecSuiv), uneCouleur, BF
            End If
            'Stockage pour le segment suivant
            unYecran = unYecSuiv
            
            For j = 2 To unNbPoints - 1
                'Calcul du point suivant pour la courbe temps/distance
                'ou la courbe vitesse/distance
                'Cumul des distances et Conversion des distances
                'des décimètres au mètre
                unD2 = unD2 + unParcours.monTabDist(j) / 10 * unParcours.monCoefEta
                'Décimètre/seconde converti en km/h
                uneV = unParcours.monTabDist(j) * unParcours.monCoefEta / unParcours.monPasMesure * 0.36
                                
                'Récupération de la couleur suivant l'appartenance à
                'telle ou telle classe de vitesses
                uneCouleur = DonnerCouleurClasseV(uneV)
                
                'Conversion en coordonnées écrans des coordonnées réelles
                'du point suivant
                unYecSuiv = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD2, uneDistMaxReelY, uneDistMaxEcranY)
                
                'Dessin de la courbe segment par segment
                's'il est entre le min et le max y écran
                'min y écran  > max y écran car les y écran orientés vers le bas en Y,
                'donc aprés conversion donnée réelle en écran le max devient < au min
                'If unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran Then
                If (unYecran <= unMinYecran And unYecran >= unMaxYecran) Or (unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran) Then
                    uneZoneDessin.Line (unX1, unYecran)-(unX2, unYecSuiv), uneCouleur, BF
                End If
                'Stockage pour le segment suivant
                unYecran = unYecSuiv
            Next j
            
            'Calcul du dernier point pour la courbe temps/distance
            'ou la courbe vitesse/distance
            'Conversion des distances des décimètres au mètre
            unD2 = unParcours.maDistPar / 10 * unParcours.monCoefEta
            If unParcours.monLastPas = 0 Then
                uneV = 0
            Else
                'Décimètre/dixième de seconde converti en km/h
                uneV = unParcours.monTabDist(unNbPoints) * unParcours.monCoefEta / unParcours.monLastPas * 3.6
            End If
            'Récupération de la couleur suivant l'appartenance à
            'telle ou telle classe de vitesses
            uneCouleur = DonnerCouleurClasseV(uneV)
            'Conversion en coordonnées écrans des coordonnées réelles
            'du point suivant
            unYecSuiv = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD2, uneDistMaxReelY, uneDistMaxEcranY)
            'Dessin de la courbe segment par segment
            's'il est entre le min et le max y écran
            'min y écran  > max y écran car les y écran orientés vers le bas en Y,
            'donc aprés conversion donnée réelle en écran le max devient < au min
            'If unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran Then
            If (unYecran <= unMinYecran And unYecran >= unMaxYecran) Or (unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran) Then
                uneZoneDessin.Line (unX1, unYecran)-(unX2, unYecSuiv), uneCouleur, BF
            End If
        End If
    Next i

    'Dessin des lignes de rappel des repères en pointillé noir
    uneZoneDessin.DrawWidth = TraitFin
    uneZoneDessin.DrawStyle = vbDashDot
    If unPrint Then
        'On fixe le x de fin de dessin pour les traits de rappel des repères
        'en impression au x max plus la longueur du libellé en X et un petit décalage
        unXFinDessin = unMaxXecran + Printer.TextWidth(unLibX) + PicBoxTop
        unDecVertical = Printer.TextHeight("W")
    Else
        'On fixe le x de fin de dessin pour les traits de rappel des repères
        'sur l'écran au x de fin de la picture box
        unXFinDessin = unMaxXecran + uneMargeD
    End If
    
    unNbRep = uneForm.maColRepere.Count
    For i = 1 To unNbRep
        'Dessin de la ligne de rappel et du nom du repère
        's'il est entre le min et le max y écran
        'min y écran  > max y écran car les y écran orientés vers le bas en Y,
        'donc aprés conversion donnée réelle en écran le max devient < au min
        Set unRep = uneForm.maColRepere(i)
        unYecran = ConvertirEnEcran(unMaxYecran, unMaxYreel - unRep.monAbsCurv, uneDistMaxReelY, uneDistMaxEcranY)
        If unYecran <= unMinYecran And unYecran >= unMaxYecran Then
            uneZoneDessin.Line (unMinXecran, unYecran)-(unXFinDessin, unYecran), QBColor(0)
            If unPrint Then
                'Impression du nom court du repère
                Printer.CurrentX = unXFinDessin
                If i = unNbRep And unNbRep > 1 Then
                    'On évite le chevauchment des noms courts
                    'deux derniers repères en impressions
                    unYprec = ConvertirEnEcran(unMaxYecran, unMaxYreel - uneForm.maColRepere(unNbRep - 1).monAbsCurv, uneDistMaxReelY, uneDistMaxEcranY)
                    If unYprec - unYecran < unDecVertical Then
                        unDecHicon = unDecVertical
                        unDecVertical = unDecVertical * 2
                    End If
                End If
                Printer.CurrentY = unYecran - unDecVertical / 2
                uneZoneDessin.Print unRep.monNomCourt
                'unXpos = unXFinDessin + unMaxWidthNomRep 'Largeur maxi de noms courts
                unXpos = unXFinDessin + Printer.TextWidth(unRep.monNomCourt) + 0.1 * UnCmEnTwips ' 0.1 cm = 1 mm
                unYpos = unYecran - unDecHicon / 2 - unRep.monIcone.Height / 4 '2
                Printer.PaintPicture unRep.monIcone.Picture, unXpos, unYpos, unRep.monIcone.Width / 2, unRep.monIcone.Height / 2
            End If
        End If
    Next i
    
    'Dessin du repère XY avec graduations principales et secondaires
    uneZoneDessin.DrawWidth = TraitFin
    uneZoneDessin.DrawStyle = vbSolid
    uneZoneDessin.Line (unMinXecran, unMinYecran)-(unMaxXecran, unMinYecran), 0
    uneZoneDessin.Line (unMinXecran, unMinYecran)-(unMinXecran, unMaxYecran), 0
    
    'Affichage du libellé sur l'axe des x
    uneZoneDessin.CurrentX = unMaxXecran
    uneZoneDessin.CurrentY = unMinYecran - uneZoneDessin.TextHeight(unLibX)
    uneZoneDessin.Print unLibX
        
    'Restauration du pointeur souris par défaut
    uneForm.MousePointer = vbDefault
End Sub

Public Sub TrouverPasGradEtModifierMinMax(unPasGrad1 As Long, unPasGrad2 As Long, unMaxReel As Single, unMinReel As Single)
    'Calcul des pas de graduations primaires et secondaires et arrondis
    'de la valeur mini à la graduation secondaire juste inférieure
    'et de la valeur maxi à la graduation secondaire juste supérieure
    Dim unNumGrad As Long, unReel As Double
       
    'Calcul du pas de graduations primaires et secondaires
    'unPasGrad1 = Int(Log(unMaxReel) / Log(10))
    'unPasGrad1 = Int(Log(unMaxReel - unMinReel) / Log(10))
    If (unMaxReel - unMinReel) < 40 Then
        'Cas où il y a moins de 40 mètres de parcours ==> unPasGrad1 = 1
        'c'est trop petit d'où plantage on met 10
        unPasGrad1 = 10
        If unMaxReel = unMinReel Then unMaxReel = unMinReel + 1
    Else
        'Cas normal
        unPasGrad1 = Log(unMaxReel - unMinReel) / Log(10) 'Arrondi par vb car long = double
        unPasGrad1 = Exp(unPasGrad1 * Log(10))
        If (unMaxReel - unMinReel) \ unPasGrad1 <= 1 Then
            'S'il n' y a qu'une seule graduation à mettre entre le min et le max
            'on descend d'un niveau de graduation en / par 10
            unPasGrad1 = unPasGrad1 / 10
        End If
    End If
    unPasGrad2 = unPasGrad1 \ 10
    
    'Mise à jour du maxi en distance pour avoir la valeur
    'arrondie à la graduation secondaire juste supérieure
    unReste = (unMaxReel Mod unPasGrad1)
    unNumGrad = unReste \ unPasGrad2
    'unNumGrad = unNumGrad - (unNumGrad > 0) - (unReste < unPasGrad2 And unReste > 0)
    unNumGrad = unNumGrad - (unReste Mod unPasGrad2 > 0)
    'Car True = -1 et False = 0 en VB
    '(Cas où la graduation juste supérieure égale le max)
    unMaxReel = (unMaxReel \ unPasGrad1) * unPasGrad1 + unNumGrad * unPasGrad2
    
    'Mise à jour du mini en distance pour avoir la valeur
    'arrondie à la graduation secondaire juste inférieure
    unNumGrad = (unMinReel Mod unPasGrad1) \ unPasGrad2
    unMinReel = (unMinReel \ unPasGrad1) * unPasGrad1 + unNumGrad * unPasGrad2
End Sub

Public Sub ArrondirXMinXMaxGrad2(unPasXGrad1 As Single, unPasXGrad2 As Single, unMaxXreel As Single, unMinXreel As Single)
    'Calcul des pas de graduations primaires et secondaires et arrondis
    'de la valeur mini à la graduation secondaire juste inférieure
    'et de la valeur maxi à la graduation secondaire juste supérieure
    Dim unPasXGrad1Tmp As Long, unPasXGrad2Tmp As Long
    Dim unMaxXreelTmp As Single, unMinXreelTmp As Single, unRes As Single
    
    unMaxXreelTmp = unMaxXreel
    unMinXreelTmp = unMinXreel
    If unMaxXreel - unMinXreel >= 10 Then
        'Arrondi à l'entier juste supérieure du maxX
        unMaxXreel = Int(unMaxXreel) + Abs(unMaxXreel - Int(unMaxXreel) > 0)
        'Arrondi à l'entier juste inférieure du minX
        unMinXreel = Int(unMinXreel)
        TrouverPasGradEtModifierMinMax unPasXGrad1Tmp, unPasXGrad2Tmp, unMaxXreel, unMinXreel
        unPasXGrad1 = unPasXGrad1Tmp
        unPasXGrad2 = unPasXGrad2Tmp
    Else
        'Cas où l'écart entre minx et maxx trop petit
        '==> graduation à un chiffre aprés virgule
        unPasXGrad1 = 1
        unPasXGrad2 = 0.2
        'Calcul du min et max avec ses nouveaux pas de graduations
        unQuotient = unMaxXreelTmp / unPasXGrad2
        'Si la division ne donne pas un résultat entier
        'on rajoute une sous-graduation
        If Abs(Int(unQuotient) - unQuotient) < Epsilon Then
            unMaxXreel = unQuotient * unPasXGrad2
        Else
            unMaxXreel = (Int(unQuotient) + 1) * unPasXGrad2
        End If
        unRes = unMinXreelTmp / unPasXGrad2
        unMinXreel = Int(unRes) * unPasXGrad2
    End If
End Sub

Public Function SelectionnerParcours(uneForm As Form, unXecran As Single, unYecran As Single) As Integer
    'Retourne l'indice dans la collection des parcours de la form se trouvant
    'sous le X et Y écran
    Dim unEpsilonX As Single, unEpsilonY As Single
    Dim unNbParcours As Integer, unNbPoints As Long
    Dim unX0 As Single, unX00 As Single
    Dim unParcours As Parcours, j As Long
    Dim unX As Single, unY As Single, uneDist As Single
    Dim unX1 As Single, unY1 As Single
    Dim unX2 As Single, unY2 As Single
    Dim unEspilonTwips As Single
    Dim uneMargeG As Single, uneMargeD As Single
    Dim uneMargeH As Single, uneMargeB As Single
    Dim uneDistM1M2 As Single
    Dim unMinXreel As Single, unMaxXreel As Single
    Dim unMaxYecran As Single, unMinXecran As Single
    
    'Affichage du sablier en pointeur souris pour symboliser l'attente
    uneForm.MousePointer = vbHourglass
    'Initialisation
    SelectionnerParcours = 0
    unEspilonTwips = 60
    'Vidage de la listbox des parcours trouvés
    frmChoixPar.Visible = False
    frmChoixPar.ListParTrouv.Clear
    
    'Conversion d'une abscisse écran de unEspilonTwips twips en abscisse réelle
    'sur OX et stockage de la picture où on dessine
    If uneForm.TabData.Tab = OngletCbeDT Then
        'Calcul des marges de travail et des distances max écran en X et en Y
        'de la picture box de la courbes DT.
        FixerMargesPicBox uneForm, uneForm.PicBoxDT, uneMargeG, uneMargeD, uneMargeH, uneMargeB
        'Récupération des minimum et maximum réel en X
        unMinXreel = uneForm.monMinT
        unMaxXreel = uneForm.monMaxT
        Set maPicBox = uneForm.PicBoxDT
        maPicBox.Tag = ""
        frmChoixPar.Tag = "DT"
    ElseIf uneForm.TabData.Tab = OngletCbeDV Then
        'Calcul des marges de travail et des distances max écran en X et en Y
        'de la picture box de la courbes DT.
        FixerMargesPicBox uneForm, uneForm.PicBoxDV, uneMargeG, uneMargeD, uneMargeH, uneMargeB
        'Récupération des minimum et maximum réel en X
        unMinXreel = uneForm.monMinV
        unMaxXreel = uneForm.monMaxV
        Set maPicBox = uneForm.PicBoxDV
        maPicBox.Tag = ""
        frmChoixPar.Tag = "DV"
        
    Else
        MsgBox MsgErreurProg + MsgErreurTypeCourbeInconnu + MsgIn + "ModuleMain:SelectionnerParcours", vbCritical
        Exit Function
    End If
    
    'Stokage du min y écran et du max X écran
    unMaxYecran = uneMargeH
    unMinXecran = uneMargeG
    
    'Conversion X écran en X réel
    unEpsilonX = DonnerDistReel(unEspilonTwips, unMaxXreel - unMinXreel, uneForm.maDistMaxEcranX)
    unX = ConvertirEnReel(unMinXreel, unXecran - uneForm.monMinXecran, unMaxXreel - unMinXreel, uneForm.maDistMaxEcranX)
    
    'Conversion d'une distance de unEspilonTwips twips en distance réelle sur OY et du Y écran
    unEpsilonY = DonnerDistReel(unEspilonTwips, uneForm.monMaxD - uneForm.monMinD, uneForm.maDistMaxEcranY)
    unY = ConvertirEnReel(uneForm.monMaxD, uneForm.monMaxYecran - unYecran, uneForm.monMaxD - uneForm.monMinD, uneForm.maDistMaxEcranY)
    
    unEpsilonXY = Sqr(unEpsilonY * unEpsilonY + unEpsilonX * unEpsilonX)
    
    'Parcours de tous les parcours utilisés
    unNbParcours = uneForm.maColParcours.Count
    If unNbParcours = 2 Then
        'Cas où il y a deux parcours dans la collection
        'le seul plus le parcours moyen on ne s'intéresse qu'au parcours unique
        i0 = 2
    Else
        i0 = 1
    End If
    
    For i = unNbParcours To i0 Step -1
        'On parcours à l'envers pour sélectionner en dernier le parcours moyen
        'car son dernier pas peut être grand, d'où un grand englobant
        'qui entraine la sélection du parcours à tous les picks vers la fin
        Set unParcours = uneForm.maColParcours(i)
        If unParcours.monIsUtil Then
            'Tout démarrer en 0,0 en coordonnées réelles les courbes DV et DT
            'ainsi le premier segment est cliquable si le premier top est donné
            'en cours de roulage.
            unX0 = 0
            uneDist = 0
            
            'Recup des tableaux de coordonnées x des points
            unNbPoints = unParcours.monNbPas
            
            For j = 1 To unNbPoints
                'Recup des coordonnées x des points
                If uneForm.TabData.Tab = OngletCbeDT Then
                    'Conversion des dixièmes de seconde et des secondes en minutes
                    If j = unNbPoints Then
                        unX00 = unParcours.monFirstPas / 600 + (unNbPoints - 2) * unParcours.monPasMesure / 60 + unParcours.monLastPas / 600
                    Else
                        unX00 = unParcours.monFirstPas / 600 + (j - 1) * unParcours.monPasMesure / 60
                    End If
                ElseIf uneForm.TabData.Tab = OngletCbeDV Then
                    'Calcul de la vitesse instantanée en km/h
                    unX00 = CalculerVitInstant(j, unParcours)
                End If
                
                'Tri pour que X1 <= X2
                If unX0 < unX00 Then
                    unX1 = unX0
                    unX2 = unX00
                Else
                    unX2 = unX0
                    unX1 = unX00
                End If
                
                'Calcul du Y
                uneDist = uneDist + unParcours.monTabDist(j - 1) * unParcours.monCoefEta / 10
                unY1 = uneDist 'Stockage pour l'incrémentation suivante
                unY2 = uneDist + unParcours.monTabDist(j) * unParcours.monCoefEta / 10
                
                'Si point confondu on ne fait rien, on passe au suivant
                unPtConfondu = (unX1 = unX2 And unY1 = unY2)
                If unPtConfondu = False Then
                    'Recherche si on a cliqué prés du segment M1(x0,y1)-M2(x00,y2)
                    'on a X1 = min(X0,X00) et X2 = max(X0,X00), pour les y pas besoins
                    'car les y = distance parcourue augmente toujours donc Y1 < Y2
                    'D'abord on regarde si X1 < X < X2, puis Y1 < Y < Y2 et enfin
                    'si la distance à la droite passant par M1 de coordonnées (unX0, unY1)
                    'et M2 de coordonnées (unX00, unY2) est < espilon en coordonnées écran
                    If (unX1 - unEpsilonX < unX) And (unX < unX2 + unEpsilonX) Then
                        If (unY1 - unEpsilonY < unY) And (unY < unY2 + unEpsilonY) Then
                            'Calcul de la distance à la droite passant par M1M2
                            'ax+by+c = 0, a = y2-y1, b=x0-x00, c = x00y1-x0y2
                            'en coordonnées écran car en écran on est en repère cartésien
                            'donc la formule de calcul de distance est bonne
                            unX1 = ConvertirEnEcran(unMinXecran, unX0 - unMinXreel, unMaxXreel - unMinXreel, uneForm.maDistMaxEcranX)
                            unY1 = ConvertirEnEcran(unMaxYecran, uneForm.monMaxD - unY1, uneForm.monMaxD - uneForm.monMinD, uneForm.maDistMaxEcranY)
                            unX2 = ConvertirEnEcran(unMinXecran, unX00 - unMinXreel, unMaxXreel - unMinXreel, uneForm.maDistMaxEcranX)
                            unY2 = ConvertirEnEcran(unMaxYecran, uneForm.monMaxD - unY2, uneForm.monMaxD - uneForm.monMinD, uneForm.maDistMaxEcranY)
                            uneDistM1M2 = (unY2 - unY1) * unXecran + (unX1 - unX2) * unYecran + (unX2 * unY1 - unX1 * unY2)
                            uneDistM1M2 = Abs(uneDistM1M2) / Sqr((unX2 - unX1) * (unX2 - unX1) + (unY2 - unY1) * (unY2 - unY1))
                            If uneDistM1M2 < unEspilonTwips Then
                                SelectionnerParcours = i
                                'Ajout dans la listbox des parcours trouvés
                                frmChoixPar.ListParTrouv.AddItem unParcours.monNom + " (" + Mid(unParcours.monJourSemaine, 1, 2) + " " + Format(unParcours.maDate) + " " + Mid(Format(unParcours.monHeureDebut), 1, 5) + ")"
                                frmChoixPar.ListParTrouv.ItemData(frmChoixPar.ListParTrouv.NewIndex) = i
                                'Sortie du for bouclant sur les points du parcours
                                'on passe au parcours suivant
                                Exit For
                            End If
                        End If
                    End If
                End If
                'Stockage pour incrément suivant
                unX0 = unX00
            Next j
        End If
    Next i
    
    'Ouverture de la fenêtre de choix du parcours à sélectionner si plusieurs
    'parcours proches du click souris
    If frmChoixPar.ListParTrouv.ListCount > 1 Then
        CentrerFenetreEcran frmChoixPar
        frmChoixPar.Show vbModal
        If maPicBox.Tag = "" Then
            'Cas où aucun parcours choisi
            '===> click sur bouton Annuler de la fenêtre choix parcours
            SelectionnerParcours = 0
        Else
            'Cas où un parcours  a été choisi
            SelectionnerParcours = CInt(maPicBox.Tag)
        End If
        'Remise à vide du tag de la fenêtre de choix
        frmChoixPar.Tag = ""
    End If
    
    'Fermeture de la fenêtre de choix du parcours à sélectionner si plusieurs
    'parcours proches du click souris, car le frmChoixPar.visible=false mis
    'au début de cette fonction alloue cette fenêtre en mémoire
    Unload frmChoixPar
    
    'Restauration du pointeur souris par défaut
    uneForm.MousePointer = vbDefault
End Function


Public Function VerifierNomCourtUnique(uneForm As Form, unNomCourt As String) As Boolean
    'Vérification de l'unicité d'un nom court dans un itinéraire (= une form)
    VerifierNomCourtUnique = True
    For i = 1 To uneForm.maColRepere.Count
        If UCase(uneForm.maColRepere(i).monNomCourt) = UCase(unNomCourt) Then
            VerifierNomCourtUnique = False
            Exit For
        End If
    Next i
End Function

Public Function DonnerLigneRepere() As Integer
    'Récup du numéro de ligne dans le spread repère de la fenêtre fille active
    'grâce à la clé d'identification du repère dont l'icône est sélectionné
    'Cette clé a été auparavant stocké dans le tag de la MDI mère
    monIti.SpreadRepere.Col = monIti.SpreadRepere.MaxCols
    For i = 1 To monIti.SpreadRepere.MaxRows
        monIti.SpreadRepere.Row = i
        If monIti.SpreadRepere.Text = monIti.Tag Then
            DonnerLigneRepere = i
            Exit For
        End If
    Next i
End Function

Public Function DonnerYRepMax(uneFrmD As frmDocument) As Long
    'Retourne le Y maxi des repères d'une form (= itinéraire)
    Dim unYRepMax As Long
    unYRepMax = -100000
    For i = 1 To uneFrmD.maColRepere.Count
        If uneFrmD.maColRepere(i).monAbsCurv > unYRepMax Then
            unYRepMax = uneFrmD.maColRepere(i).monAbsCurv
        End If
    Next i
    DonnerYRepMax = unYRepMax
End Function

Public Function DonnerYRepMin(uneFrmD As frmDocument) As Long
    'Retourne le Y mini des repères d'une form (= itinéraire)
    Dim unYRepMin As Long
    unYRepMin = 10000000
    For i = 1 To uneFrmD.maColRepere.Count
        If uneFrmD.maColRepere(i).monAbsCurv < unYRepMin Then
            unYRepMin = uneFrmD.maColRepere(i).monAbsCurv
        End If
    Next i
    DonnerYRepMin = unYRepMin
End Function


Public Function DonnerValGrad(uneFrmD As Form, unLong As Long, unRnd As Byte) As Long
    'Fonction retournant la graduation de niveau juste supérieure, si unRnd = 1
    'ou juste inférieure si unRnd = 0
    'à la valeur passé en paramètre d'une fenêtre itinéraire
    
    Dim unQuot1 As Long, unReste1 As Long
    Dim unQuot2 As Long, unReste2 As Long
    
    If unRnd > 1 Then
        MsgBox "Erreur de programmation dans DonnerValGrad : unRnd vaut 0 ou 1"
        Exit Function
    End If
    'Code pour les tests unitaires de cette fonction
    'uneRep = vbRetry
    'While uneRep = vbRetry
        'unechaine = InputBox("Entrer le nombre à arrondir à la graduation inférieure ou supérieure, ainsi que les deux niveaux de graduations", "Test unitaire", "1111-100-10-0")
        'unLong = Val(unechaine)
        'uneFrmD.monPasGrad1 = Val(Mid(unechaine, Len(Format(unLong)) + 2))
        'uneFrmD.monPasGrad2 = Val(Mid(unechaine, Len(Format(unLong)) + Len(Format(uneFrmD.monPasGrad1)) + 3))
        'If Mid(unechaine, Len(unechaine)) = "0" Then
        '    unRnd = 0
        'Else
        '    unRnd = 1
        'End If
    
    unQuot1 = unLong \ uneFrmD.monPasGrad1
    unReste1 = unLong Mod uneFrmD.monPasGrad1
    
    unQuot2 = unReste1 \ uneFrmD.monPasGrad2
    unReste2 = unReste1 Mod uneFrmD.monPasGrad2
    
    If unReste2 > 0 Then unQuot2 = unQuot2 + unRnd
    DonnerValGrad = unQuot1 * uneFrmD.monPasGrad1 + unQuot2 * uneFrmD.monPasGrad2
        'uneRep = MsgBox("Valeur ini = " + Format(unLong) + " Pas1 = " + Format(uneFrmD.monPasGrad1) + " Pas2 = " + Format(uneFrmD.monPasGrad2) + " ===> Valeur Arrondi sup = " + Format(DonnerValGrad), vbRetryCancel)
    'Wend
End Function

Public Sub ViderColParcours(uneColPar As ColParcours)
    'Procédure vidant une collection de parcours
    For i = 1 To uneColPar.Count
        uneColPar.Remove 1
    Next i
End Sub

Public Sub ViderColRepere(uneColRep As ColRepere)
    'Procédure vidant une collection de repères
    For i = 1 To uneColRep.Count
        uneColRep.Remove 1
    Next i
End Sub

Public Sub ViderCollection(uneCol As Collection)
    'Procédure vidant une collection
    For i = 1 To uneCol.Count
        uneCol.Remove 1
    Next i
End Sub

Public Sub DonnerMaxDistDureeVit(uneColParcours As ColParcours, uneDistMax As Single, uneDureeMax As Single, uneVitMax As Single)
    'Procédure donnant la distance maxi parcourue, la durée de parcours maxi et la vitesse maxi instantanée
    'dans une liste de parcours en tenant compte uniquement des parcours utilisés
    'Les résultats sont stockés dans les trois paramètres de type single
    Dim unPar As Parcours, uneDist As Single, uneDuree As Single
    Dim uneVit As Single, j As Long
    
    uneDistMax = 0
    uneDureeMax = 0
    uneVitMax = 0
    
    For i = 1 To uneColParcours.Count
        Set unPar = uneColParcours(i)
        If unPar.monIsUtil Then
            'Initialisation des champs vitesses du parcours
            unPar.maVmax = 0
            unPar.maVmin = 0
            unPar.maVmoy = 0
            'Conversion des décimétres en m
            uneDist = unPar.maDistPar / 10 * unPar.monCoefEta
            'Stockage de la distance parcourue maxi
            If uneDist > uneDistMax Then uneDistMax = uneDist
            'Conversion des dixièmes de secondes en minutes
            uneDuree = unPar.maDuree / 600
            'Stockage de la durée de parcours maxi
            If uneDuree > uneDureeMax Then uneDureeMax = uneDuree
            
            'Calcul de la vitesse maxi instantanée en km/h du parcours
            'Calcul pour le premier pas de mesure en km/h
            If unPar.monFirstPas = 0 Then
                uneVit = 0
            Else
                uneVit = unPar.monTabDist(1) * unPar.monCoefEta / unPar.monFirstPas * 3.6
            End If
            If uneVit > uneVitMax Then uneVitMax = uneVit
            'Calcul pour les pas situés entre le deuxième et l'avant dernier
            For j = 2 To unPar.monNbPas - 1
                uneVit = unPar.monTabDist(j) * unPar.monCoefEta / unPar.monPasMesure * 0.36
                If uneVit > uneVitMax Then uneVitMax = uneVit
            Next j
            'Calcul pour le dernier pas de mesure en km/h
            If unPar.monLastPas = 0 Then
                uneVit = 0
            Else
                uneVit = unPar.monTabDist(unPar.monNbPas) * unPar.monCoefEta / unPar.monLastPas * 3.6
            End If
            If uneVit > uneVitMax Then uneVitMax = uneVit
        End If
    Next i
    
    'Arrondi éventuel à l'entier juste supérieure
    If uneDistMax > Int(uneDistMax) Then uneDistMax = Int(uneDistMax) + 1
    If uneDureeMax > Int(uneDureeMax) Then uneDureeMax = Int(uneDureeMax) + 1
    If uneVitMax > Int(uneVitMax) Then uneVitMax = Int(uneVitMax) + 1
End Sub

Public Sub DonnerMaxDistDuree(uneColParcours As ColParcours, uneDistMax As Single, uneDureeMax As Single)
    'Procédure donnant la distance maxi parcourue et la durée de parcours maxi
    'dans une liste de parcours en tenant compte uniquement des parcours utilisés
    'Les résultats sont stockés dans les deux paramètres de type single
    Dim unPar As Parcours, uneDist As Single, uneDuree As Single
    
    uneDistMax = 0
    uneDureeMax = 0
    
    For i = 1 To uneColParcours.Count
        Set unPar = uneColParcours(i)
        If unPar.monIsUtil Then
            'Conversion des décimétres en m
            uneDist = unPar.maDistPar / 10 * unPar.monCoefEta
            'Stockage de la distance parcourue maxi
            If uneDist > uneDistMax Then uneDistMax = uneDist
            'Conversion des dixièmes de secondes en minutes
            uneDuree = unPar.maDuree / 600
            'Stockage de la durée de parcours maxi
            If uneDuree > uneDureeMax Then uneDureeMax = uneDuree
        End If
    Next i
    'Arrondi de la dist max au décimtre supérieure à cause d'un problème
    'd'arrondi entre les doubles et les singles en VB
    uneDistMax = (Int(uneDistMax * 10) + 1) / 10
End Sub

Public Sub ActualiserParcoursMoyen(unParMoyen As Parcours, uneColParcours As ColParcours, uneD1 As Long, uneD2 As Long)
    'Actualisation d'un parcours moyen à partir d'une liste de parcours
    'en ne prenant en compte que ceux utilisés
    'et retourne la valeur de la vitesse maxi instantanée
    Dim unPar As Parcours, uneColIndexUtil As New Collection
    Dim uneColNbVal As New Collection, uneDistMin As Single
    Dim unNbRep As Integer, unNbRepTop As Long
    Dim unTK As Single, unTK_1 As Single, unEpsilon As Byte
    Dim unT As Single, unPas As Long, uneDist As Long, unPasDist As Long
    Dim unN As Long, unN0 As Long, k As Long, uneDist0 As Long
    Dim unNbPas As Long, i As Long, unFirstPas As Single
    Dim unTabTempsRep As Variant, unTabAbsRep As Variant, uneDtmp As Long
    Dim unTabNbVal(1 To NbPasMax) As Byte, uneInterDist As Single 'Long
    Dim unTabTmpPar(0 To NbPasMax) As Single, unNbPasDist As Long
    
    'Affectation par défaut avec la date du jour d'utilisation de MiTemps
    unParMoyen.monNom = "Parcours moyen"
    unParMoyen.maDate = Date
    unParMoyen.monJourSemaine = DonnerJourSemaine(Date)
    unParMoyen.maCouleur = QBColor(0) ' = noir
    unParMoyen.monCoefEta = 1 'ne sert pas pour le parcours moyen
    
    'Boucle pour trouver les distance parcourue et durée totales moyennes
    'et stockage dans une collection des index de parcours utilisés et déterminé
    'le nombre de pas de mesure de 1 secondes pour dimensionner le tableau des
    'distances du parcours moyen. on recherche aussi la longueur de parcours maxi
    uneDistMin = 2000000000
    'On met au maxi des long soit deux milliards de décimètre, soit 200 000 km
    unParMoyen.maDistPar = 0
    unParMoyen.maDuree = 0
    unParMoyen.monPasMesure = 0
    unParMoyen.maVmoy = 0
    For i = 2 To uneColParcours.Count
        'De 2 à n car le premier de la collection est tjs le parcours moyen
        Set unPar = uneColParcours(i)
        If unPar.monIsUtil Then
            uneColIndexUtil.Add i
            If unPar.maDistPar * unPar.monCoefEta < uneDistMin Then
                uneDistMin = unPar.maDistPar * unPar.monCoefEta
            End If
            'Ici la vitesse moyenne sera en décimètre par dixième de seconde
            unParMoyen.maVmoy = unParMoyen.maVmoy + unPar.maDistPar * unPar.monCoefEta / unPar.maDuree
            unParMoyen.monPasMesure = unParMoyen.monPasMesure + unPar.monPasMesure
        End If
    Next i
    
    'Si aucun parcours utilisé ==> on sort
    If uneColIndexUtil.Count = 0 Then
        unParMoyen.monNbPas = 1
        unParMoyen.monTabDist(1) = unParMoyen.maDistPar
        Exit Sub
    End If
    
    'Finalisation des calculs des moyennes en divisant par le nb de parcours utilisés
    'en arrondissant, pour la distance on prend la distmin sinon les algo du parcours
    'moyen donne des choses bizarres pour la courbe distance/vitesse
    unParMoyen.maDistPar = Int(uneDistMin)
    'Calcul et Conversion de la vitesse moyenne en décimètre par dixième de seconde en km/h
    unParMoyen.maVmoy = unParMoyen.maVmoy / uneColIndexUtil.Count * 3.6
    'Pour le pas de mesure l'unité est la seconde
    unParMoyen.monPasMesure = Int(unParMoyen.monPasMesure / uneColIndexUtil.Count)
    
    'Calcul du pas de discrétisation pour calculer les temps de passage en décimètre
    '= distance en décimètre parcourue en pas mesure moyen à la vitesse moyenne en km/h
    unPasDist = Int(unParMoyen.maVmoy / 3.6 * 10 * unParMoyen.monPasMesure)
    'Affectation des valeurs du premier pas de mesure
    'On choisit celui du premier parcours utilisé de la collection de parcours
    'Ainsi, si un seul parcours utilisé le parcours moyen = l'unique parcours utilisé
    unParMoyen.monFirstPas = uneColParcours(uneColIndexUtil(1)).monFirstPas
    
    
    'Affichage de la fenêtre de progression du calcul
    If unParMoyen.maDistPar Mod unPasDist = 0 Then
        unNbPasDist = unParMoyen.maDistPar \ unPasDist
    Else
        unNbPasDist = unParMoyen.maDistPar \ unPasDist + 1
    End If
    uneVal100 = uneColIndexUtil.Count * unNbPasDist 'valeur du 100%
    unPasProg = 10
    unNbCalc = 0
    frmWaitCalcul.Show
    frmWaitCalcul.Caption = "Progression du calcul du parcours moyen"
    frmWaitCalcul.LabelN.Left = frmWaitCalcul.ProgressBar1.Left
    frmWaitCalcul.LabelN.Caption = "0%"
    
    'Calcul du tableau des temps moyens de parcours
    For j = 1 To uneColIndexUtil.Count
        'Calcul du pas pour chaque parcours utilisé
        Set unPar = uneColParcours(uneColIndexUtil(j))
        uneDist = 0
        unN = unPar.monNbPas
        unN0 = 1
        uneDistCumul = 0
        For i = 1 To unNbPasDist
            'Affectation du nombre de calcul pour calculer la progression
            'tous les unPasProg % effectués
            unNbCalc = unNbCalc + 1
            If unNbCalc Mod (uneVal100 \ unPasProg) = 0 Then
                frmWaitCalcul.Show
                CentrerFenetreEcran frmWaitCalcul
                frmWaitCalcul.ProgressBar1.Value = CLng(unNbCalc / uneVal100 * 100)
                frmWaitCalcul.LabelN.Caption = Format(frmWaitCalcul.ProgressBar1.Value) + " %"
                frmWaitCalcul.LabelN.Left = frmWaitCalcul.ProgressBar1.Left + frmWaitCalcul.ProgressBar1.Width * frmWaitCalcul.ProgressBar1.Value / 100
            End If
            
            'Stockage de la distance cumulé au pas en distance précédent
            uneDist0 = uneDist
            'Calcul des temps de parcours pour une distance donnée
            If uneDist + unPasDist > unParMoyen.maDistPar Then
                'Cas où l'on dépasse la valeur moyenne, on s'y ramène
                'donc unedist est tjs < ou = à la distance moyenne
                uneDist = unParMoyen.maDistPar
            Else
                uneDist = uneDist + unPasDist
            End If
            If uneDist < unParMoyen.maDistPar + EpsilonDist Then
                'Cas où la distance ne dépasse pas la distance du parcours j
                '==> Prise en compte pour la valeur moyenne
                'Calcul du temps de passage à la distance uneDist
                For k = unN0 To unN
                    uneDistCumul = uneDistCumul + unPar.monTabDist(k) * unPar.monCoefEta
                    If uneDistCumul > uneDist - EpsilonDist Then
                        'Calcul du temps de passage à ce pas là pour le parcours en cours
                        uneDistCumul0 = uneDistCumul - unPar.monTabDist(k) * unPar.monCoefEta
                        'Stockage du pas où l'on dépasse la distance uneDist
                        'ou que l'on égalise cette distance pour le calcul au prochain pas
                        unN0 = k
                        If k = 1 Then
                            If unPar.monTabDist(1) = 0 Then
                                unT = 0
                            Else
                                unT = unPar.monFirstPas * (uneDist - uneDistCumul0) / unPar.monTabDist(1) / unPar.monCoefEta
                            End If
                        ElseIf k = unN Then
                            If unPar.monTabDist(unN) = 0 Then
                                unT = unPar.monFirstPas + (k - 2) * unPar.monPasMesure * 10
                            Else
                                unT = unPar.monFirstPas + (k - 2) * unPar.monPasMesure * 10 + (uneDist - uneDistCumul0) / unPar.monTabDist(unN) / unPar.monCoefEta * unPar.monLastPas
                            End If
                        Else
                            unT = unPar.monFirstPas + (k - 2 + (uneDist - uneDistCumul0) / unPar.monTabDist(k) / unPar.monCoefEta) * unPar.monPasMesure * 10
                        End If
                        'Remise de la distance cumulé
                        'à celle du pas précédent pour le calcul au prochain pas
                        uneDistCumul = uneDistCumul0
                        Exit For 'on sort du for k et on passe au pas en distance suivant
                    End If
                Next k
                '==> Calcul de la moyenne grâce aux valeurs par pas des parcours précédent
                unTabTmpPar(i) = (unTabTmpPar(i) * unTabNbVal(i) + unT) / (unTabNbVal(i) + 1)
                '==> Prise en compte pour la valeur moyenne
                unTabNbVal(i) = unTabNbVal(i) + 1
            Else
                Exit For 'on sort du for i et on passe au parcours suivant éventuel
            End If
        Next i
    Next j

    'Calcul de la durée du parcours moyen qui correspond au temps de passage
    'à la distance moyenne, donc au dernier pas en distance arrondi au dixième
    'de seconde supérieure
    unParMoyen.maDuree = Int(unTabTmpPar(unNbPasDist)) + 1
    unTabTmpPar(0) = 0
    
    'Calcul de la distance parcourue par le parcours moyen à ce temps de passage
    'qui devient la nouvelle distance parcourue
    uneDistTmp! = uneDist0 + (unParMoyen.maDuree - unTabTmpPar(unNbPasDist - 1)) / (unTabTmpPar(unNbPasDist) - unTabTmpPar(unNbPasDist - 1)) * (unParMoyen.maDistPar - uneDist0)
    unParMoyen.maDistPar = uneDistTmp!
    'On met à jour le dernier temps de passage correspondant à la distance parcourue par
    'le parcours moyen
    unTabTmpPar(unNbPasDist) = unParMoyen.maDuree
    
    'Calcul du nombre de pas de mesure pour le parcours moyen
    'et détermination de la valeur du dernier pas de mesure
    unPas = unParMoyen.monPasMesure * 10
    'Pour être en 1/10 de secondes car la durée, firstpas et lastpas sont en 1/10 de secondes
    unNbPas = (unParMoyen.maDuree - unParMoyen.monFirstPas) \ unPas + 1
    '\ = division entière et il faut rajouter le premier pas de mesure
    If (unParMoyen.maDuree - unParMoyen.monFirstPas) Mod unPas = 0 Then
        'Cas d'un nombre entier de pas
        '==> valeur du dernier pas de mesure = le pas
        unParMoyen.monLastPas = unPas
    Else
        'Cas d'un nombre entier de pas ne suffit pas
        '==> Il faut rajouter le dernier pas de mesure
        unNbPas = unNbPas + 1
        'Et la valeur du dernier pas de mesure = le reste de la division
        'de la durée moins le premier pas, par le pas plus écart
        'entre durée moyen et min
        unParMoyen.monLastPas = unParMoyen.maDuree - (unNbPas - 2) * unPas - unParMoyen.monFirstPas
     End If
    'Récup du nombre de pas d'interdistances
    unParMoyen.monNbPas = unNbPas
    
    'Calcul des distances pour chaque pas d'une seconde pour remplir le parcours moyen
    unT = unParMoyen.monFirstPas
    k = 1
    i = 1
    unPasTmp = unPasDist 'Le pas en distance entre 1 et N-2 vaut unPasDist
    unTK = unTabTmpPar(1)
    unTK_1 = unTabTmpPar(0)
    While i < unNbPas And unParMoyen.monTabDist(i - 1) < unParMoyen.maDistPar + EpsilonDist
        'On doit remplir unParMoyen.monTabDist de 1 jusqu'à unNbPas-1
        'On fait le dernier pas après
        If unT < unTK + EpsilonDist Then
            unParMoyen.monTabDist(i) = (unT - unTK_1) / (unTK - unTK_1) * unPasTmp + (k - 1) * unPasDist
            'TabTmpPar(k) est tjs > TabTmpPar(k-1) car on est au pas de distance supérieur
            'donc le temps de parcours a forcément augmenté
            
            'Incrémentation suivante
            i = i + 1
            unT = unT + unParMoyen.monPasMesure * 10
            'Le pas de mesure est en seconde et les temps en dixième de seconde
        Else
            If k < unNbPasDist - 1 Then
                k = k + 1
                unTK = unTabTmpPar(k)
                unTK_1 = unTabTmpPar(k - 1)
            Else
                'Cas où l'on reste dans le dernier intervalle de temps parcouru
                'on doit remplir unParMoyen.monTabDist de 1 jusqu'à unNbPas-1
                k = unNbPasDist
                unPasTmp = unParMoyen.maDistPar - uneDist0
                'Le pas entre 1 et N-2 vaut la différence entre les 2 derniers distances
                unTK = unTabTmpPar(k)
                unTK_1 = unTabTmpPar(k - 1)
            End If
        End If
    Wend
    
    'Correction des pas entre la sortie de la boucle si on dépasse la distance parcourue
    'et le dernier pas théorique unNbPas, on met distance parcourue pour tous ces pas
    'De plus si on sort car i = unNbPas on remplit aussi le dernier pas avec la bonne
    'distance (cf le if ci-dessous)
    If i = unNbPas Then i = unNbPas + 1
    For k = i - 1 To unNbPas
        unParMoyen.monTabDist(k) = unParMoyen.maDistPar
    Next k
        
    'Mise à jour du tableau des distances en inter-distances comme les parcours mesurés
    For i = unNbPas To 2 Step -1
        unParMoyen.monTabDist(i) = unParMoyen.monTabDist(i) - unParMoyen.monTabDist(i - 1)
    Next i
            
    'Calcul des vitesses min, max et moyenne et de la durée, du nombre
    'et du temps d'arrêts sur le parcours total
    If uneD1 > uneD2 Then
        uneDtmp = uneD1
        uneD1 = uneD2
        uneD2 = uneDtmp
    End If
    'Les distances sont en décimètres
    unParMoyen.CalculerLesVitDistDureeEtArrets uneD1 * 10, uneD2 * 10
    
    'Calcul des abs curv et des temps de passage moyen aux repères moyens
    'Allocation dynamique des tableaux liés aux repères topés
    unTabTempsRep = unParMoyen.monTabTempsRep
    unTabAbsRep = unParMoyen.monTabAbsRep
    unNbRep = UBound(unParMoyen.monTabTempsRep)
    'Mise à zéro du nb de valeurs intervenants dans le calcul de moyenne
    'et du tableau des temps de passage
    For i = 1 To unNbRep
        unTabNbVal(i) = 0
        unTabTempsRep(i) = 0
        unTabAbsRep(i) = 0
    Next i
    For i = 1 To uneColIndexUtil.Count
        Set unPar = uneColParcours(uneColIndexUtil(i))
        'Recup du nb de tops du parcours
        unNbRepTop = UBound(unPar.monTabAbsRep)
        'Le nb de top d'un parcours = au nb de top du parcours moyen
        'si on a crée l'itinéraire à partir d'une campagne de mesure
        'et >= si on importe le parcours d'une campagne de mesure dans
        'l'itinéraire ainsi on sort des bornes des tableaux monTabTempsRep
        'et monTabAbsRep
        For j = 1 To unNbRepTop
            unTabNbVal(j) = unTabNbVal(j) + 1
            unTabTempsRep(j) = unTabTempsRep(j) + unPar.monTabTempsRep(j)
            unTabAbsRep(j) = unTabAbsRep(j) + unPar.monTabAbsRep(j) * unPar.monCoefEta
        Next j
    Next i
    For j = 1 To unNbRep
        'unTabTempsRep(j) = unTabTempsRep(j) / uneColIndexUtil.Count
        If unTabNbVal(j) = 0 Then
            unTabTempsRep(j) = 0
            unTabAbsRep(j) = unParMoyen.maDistPar '+ mesOptions.monEcartMax * (2*j + 1)
            'ecart max converti en décimètre et * 2 pour ne pas créé d'autres
            'double top abs1-abs2 < ecartmax*10
        Else
            unTabTempsRep(j) = unTabTempsRep(j) / unTabNbVal(j)
            unTabAbsRep(j) = unTabAbsRep(j) / unTabNbVal(j)
        End If
    Next j
    'Affectation des pointeurs sur les tableaux du parcours
    unParMoyen.monTabTempsRep = unTabTempsRep
    unParMoyen.monTabAbsRep = unTabAbsRep
    
    'Calcul du nombre et de la durée des double tops
    'entre deux abs curv englobant tout le parcours
    'Les distances sont en décimètres
    unParMoyen.CalculerNbEtDureeDoubleTop uneD1 * 10, uneD2 * 10
    
    'Fermeture de la fenêtre de progression du calcul
    frmWaitCalcul.Hide
    Unload frmWaitCalcul
    ViderCollection uneColIndexUtil
    Set uneColIndexUtil = Nothing
End Sub

Public Function CalculerVitInstant(unI As Long, unPar As Parcours) As Single
    'Fonction donnant la vitesse instantanée d'un
    'parcours dans le pas d'inter-distance unI en km/h
    If unI = 1 Then
        If unPar.monFirstPas = 0 Then
            CalculerVitInstant = 0
        Else
            'Conversion deci m / dixième de sec en km/h
            CalculerVitInstant = unPar.monTabDist(unI) * unPar.monCoefEta / unPar.monFirstPas * 3.6
        End If
    ElseIf unI = unPar.monNbPas Then
        If unPar.monLastPas = 0 Then
            CalculerVitInstant = 0
        Else
            'Conversion deci m / dixième de sec en km/h
            CalculerVitInstant = unPar.monTabDist(unI) * unPar.monCoefEta / unPar.monLastPas * 3.6
        End If
    Else
        'Conversion deci m/sec en km/h
        CalculerVitInstant = unPar.monTabDist(unI) * unPar.monCoefEta / unPar.monPasMesure * 0.36
    End If
End Function

Public Function FormatterTempsEnHMNS(uneDuree As Long) As String
    'Fonction retournant une chaine de caractère formattant une durée
    'en millisecondes en un format 00h 00mn 00s
    
    'On stocke la vraie durée au cas où elle est < 60 secondes
    'pour apparaitre les dixièmes de secondes
    If uneDuree < 600 And uneDuree > 0 Then
        uneDureeTmp = uneDuree
    End If
    'On vire les millisecondes de la durée
    uneDuree = CLng(uneDuree / 10)
    unNbHeure = uneDuree \ 3600
    unNbMin = (uneDuree Mod 3600) \ 60
    unNbSec = (uneDuree Mod 3600) Mod 60
    uneStringDuree = ""
    If unNbHeure > 0 Then uneStringDuree = Format(unNbHeure) + "h "
    If unNbMin > 0 Or unNbHeure > 0 Then uneStringDuree = uneStringDuree + Format(unNbMin) + "mn "
    If uneStringDuree = "" Then
        If uneDureeTmp < 600 And uneDureeTmp > 0 Then
            uneStringDuree = uneStringDuree + Format(uneDureeTmp / 10, "fixed") + "s"
        Else
            uneStringDuree = uneStringDuree + Format(unNbSec) + "s"
        End If
    ElseIf unNbSec < 10 And unNbSec > 0 Then
        uneStringDuree = uneStringDuree + "0" + Format(unNbSec) + "s"
    Else
        uneStringDuree = uneStringDuree + Format(unNbSec) + "s"
    End If
            
    'Valeur de retour
    FormatterTempsEnHMNS = uneStringDuree
End Function


Public Function DonnerCouleurClasseV(uneV As Single) As Long
    'Fonction retournant la couleur de la classe de vitesses dans
    'laquelle se trouve la vitesse uneV
    If uneV <= mesOptions.maValClasV1 Then
        DonnerCouleurClasseV = mesOptions.maCouleurClasV1
    ElseIf uneV <= mesOptions.maValClasV2 Then
        DonnerCouleurClasseV = mesOptions.maCouleurClasV2
    ElseIf uneV <= mesOptions.maValClasV3 Then
        DonnerCouleurClasseV = mesOptions.maCouleurClasV3
    ElseIf uneV <= mesOptions.maValClasV4 Then
        DonnerCouleurClasseV = mesOptions.maCouleurClasV4
    ElseIf uneV <= mesOptions.maValClasV5 Then
        DonnerCouleurClasseV = mesOptions.maCouleurClasV5
    Else
        'Cas d'une vitesse > à Classe V5
        DonnerCouleurClasseV = mesOptions.maCouleurClasV6
    End If
End Function

Public Sub DessinerHistoV(uneForm As Form)
    'Dessin des histogrammes de vitesses dans l'onglet Histogramme des vitesses
    Dim unPar As Parcours, unTabClasV(1 To 6) As Single
    Dim uneD As Single, uneV As Single, unNumClasV As Byte
    Dim unNbVitTot As Long, unY1 As Long, unY2 As Long
    Dim unYTmp As Long, j As Long, unNbVal As Long
    
    'Affichage du sablier en pointeur souris pour symboliser l'attente
    uneForm.MousePointer = vbHourglass
    
    'Test si on met moins de 10 parcours sur l'histogramme
    unNbParUtil = DonnerNbParcoursUtil(uneForm)
    If unNbParUtil > 10 Then
        MsgBox "Impossible d'afficher ou d'imprimer plus de 10 parcours dans l'histogramme des vitesses. Diminuer votre nombre de parcours utilisés.", vbExclamation
        uneForm.MousePointer = vbDefault
        unNbParUtil = 10
        'uneForm.MSChart1.Visible = False
        'Exit Sub
    Else
        'uneForm.MSChart1.Visible = True
    End If
    
    'Récup des bornes de la section de travail éventuelle
    If uneForm.CheckSection.Value = 0 Then
        'Pas de section de tavail ==> Stockage des abs début et fin du parcours
        unY1 = -100
        unY2 = 1000000
    Else
        'Stockage des abs début et fin de la section de travail du parcours
        unY1 = uneForm.maColRepere(uneForm.ComboRepDebSec.ListIndex + 1).monAbsCurv
        unY2 = uneForm.maColRepere(uneForm.ComboRepFinSec.ListIndex + 1).monAbsCurv
        If unY1 > unY2 Then
            'Pour avoir toujours Y1 <= Y2
            unYTmp = unY2
            unY2 = unY1
            unY1 = unYTmp
        End If
    End If
    
    'Calcul des vitesses et répartition dans les classes
    'et alimentation du MSChart
    uneForm.MSChart1.ColumnCount = unNbParUtil
    
    unNbPar = 0
    i = 1
    'For i = 1 To uneForm.maColParcours.Count
    While i <= uneForm.maColParcours.Count And unNbPar < unNbParUtil
        Set unPar = uneForm.maColParcours(i)
        'Initialisation du Nombre de valeurs totales situées
        'entre les bornes de la section de travail
        unNbVal = 0
        If unPar.monIsUtil Then
            unNbPar = unNbPar + 1
            uneForm.MSChart1.Plot.SeriesCollection(unNbPar).LegendText = unPar.monNom + " (" + UCase(Mid(unPar.monJourSemaine, 1, 2)) + " " + Format(unPar.maDate, "dd/mm/yy") + " " + Mid(unPar.monHeureDebut, 1, 5) + ")   "
            With uneForm.MSChart1.Plot.SeriesCollection.Item(unNbPar).DataPoints(-1)
                'Attribue la couleur du parcours au point de données.
                .Brush.Style = VtBrushStyleSolid
                'Associe la couleur parcours, transformation d'un entier long
                'en composante RGB
                unBlue = unPar.maCouleur \ CarreDe256
                unGreen = (unPar.maCouleur Mod CarreDe256) \ 256
                unRed = unPar.maCouleur - unBlue * CarreDe256 - unGreen * 256
                .Brush.FillColor.Set unRed, unGreen, unBlue
            End With
            
            '*******************************************************
            'Calcul du nombre de vitesses dans chaque classe
            '*******************************************************
            unNbVitTot = unPar.monNbPas
            'Remise à zéro du tableau des classes de vitesses
            For k = 1 To 6
                unTabClasV(k) = 0
            Next k
                
            'Calcul de la première vitesse
            'Conversion des distances des décimètres au mètre
            uneD = unPar.monTabDist(1) / 10 * unPar.monCoefEta
            If uneD >= unY1 And uneD <= unY2 Then
                'Cas où on est entre les bornes de section de travail
                If unPar.monFirstPas = 0 Then
                    uneV = 0
                Else
                    'mètre/dixième de seconde converti en km/h
                    uneV = uneD / unPar.monFirstPas * 36
                End If
                'Récupération de la classe de vitesse contenant cette vitesse
                unNumClasV = DonnerNumClasseV(uneV)
                'Incrémentation du nombre de vitesse trouvée dans cette classe
                unTabClasV(unNumClasV) = unTabClasV(unNumClasV) + 1
                'Incrémentation du nombre de vitesses classés
                unNbVal = unNbVal + 1
            End If
            
            For j = 2 To unNbVitTot - 1
                'Calcul des vitesses suivantes
                'Cumul des distances et Conversion des distances
                'des décimètres au mètre
                uneD = uneD + unPar.monTabDist(j) / 10 * unPar.monCoefEta
                If uneD >= unY1 And uneD <= unY2 Then
                    'Cas où on est entre les bornes de section de travail
                    'Décimètre/seconde converti en km/h
                    uneV = unPar.monTabDist(j) * unPar.monCoefEta / unPar.monPasMesure * 0.36
                    'Récupération de la classe de vitesse contenant cette vitesse
                    unNumClasV = DonnerNumClasseV(uneV)
                    'Incrémentation du nombre de vitesse trouvée dans cette classe
                    unTabClasV(unNumClasV) = unTabClasV(unNumClasV) + 1
                    'Incrémentation du nombre de vitesses classés
                    unNbVal = unNbVal + 1
                End If
            Next j
            
            'Calcul de la dernière vitesse
            'Conversion des distances des décimètres au mètre
            uneD = unPar.maDistPar / 10 * unPar.monCoefEta
            If uneD >= unY1 And uneD <= unY2 Then
                'Cas où on est entre les bornes de section de travail
                If unPar.monLastPas = 0 Then
                    uneV = 0
                Else
                    'Décimètre/dixième de seconde converti en km/h
                    uneV = unPar.monTabDist(unNbVitTot) * unPar.monCoefEta / unPar.monLastPas * 3.6
                End If
                'Récupération de la classe de vitesse contenant cette vitesse
                unNumClasV = DonnerNumClasseV(uneV)
                'Incrémentation du nombre de vitesse trouvée dans cette classe
                unTabClasV(unNumClasV) = unTabClasV(unNumClasV) + 1
                'Incrémentation du nombre de vitesses classés
                unNbVal = unNbVal + 1
            End If
            
            'Alimentation du MSCHART avec les % de vitesses par classes
            If unNbVal = 0 Then unNbVal = 1
            For k = 1 To 6 'Il y a 6 classes
                uneForm.MSChart1.Row = k
                uneForm.MSChart1.Column = unNbPar
                'Calcul du % de vitesses dans la classe i
                'et insertion dans le MSChart
                uneForm.MSChart1.Data = unTabClasV(k) / unNbVal * 100
            Next k
        End If
        i = i + 1
    Wend
    'Next i
    
    'Restauration du pointeur souris
    uneForm.MousePointer = vbDefault
End Sub

Public Function DonnerNumClasseV(uneV As Single) As Byte
    'Fonction retournant l'indice de la classe contenant la vitesse uneV
    'en fonction des classes de vitesses des options du logiciel
    If uneV <= mesOptions.maValClasV1 Then
        DonnerNumClasseV = 1
    ElseIf uneV <= mesOptions.maValClasV2 Then
        DonnerNumClasseV = 2
    ElseIf uneV <= mesOptions.maValClasV3 Then
        DonnerNumClasseV = 3
    ElseIf uneV <= mesOptions.maValClasV4 Then
        DonnerNumClasseV = 4
    ElseIf uneV <= mesOptions.maValClasV5 Then
        DonnerNumClasseV = 5
    Else
        DonnerNumClasseV = 6
    End If
End Function

Public Function DonnerNbParcoursUtil(uneForm As Form) As Integer
    'Fonction retournant le nombre de parcours utilisés dans la fenêtre
    'itinéraire passée en paramètres
    DonnerNbParcoursUtil = 0
    For i = 1 To uneForm.maColParcours.Count
        If uneForm.maColParcours(i).monIsUtil Then DonnerNbParcoursUtil = DonnerNbParcoursUtil + 1
    Next i
End Function

Public Sub IndiquerToutRedessiner(uneForm As Form, Optional unIdeb As Integer = 1, Optional unIfin As Integer = 6)
    'Initialisation des indicateurs de redessin des onglets de unIdeb à unIfin
    'à vrai pour déclencher le dessin lors de leur activation
    Dim i As Integer
    
    For i = unIdeb To unIfin
        uneForm.SetTabRedOnglet i, True
    Next i
End Sub

Public Sub DessinerDesParcours(uneZoneDes As PictureBox, uneColPar As ColParcours, uneMargeD As Single, uneMargeG As Single, uneMargeB As Single, uneMargeH As Single, unMinXreel As Single, unMaxXreel As Single, unMinYreel As Single, unMaxYreel As Single, Optional unIndParChoisi As Integer = 0)
    'Dessin de la courbe Distance/Temps et des repères topés d'une collection de parcours dans
    'une picture box avec respect des marges entre un min en X et un max en X réel
    'et entre un min en Y réel et un max Y réel
    'Donc on mappe Largeur picturebox - marge droite - marge gauche sur maxXreel-minXreel
    'et Hauteur picturebox - marge haut - marge bas sur maxYreel-minYreel
    Dim unParcours As Parcours, j As Long, unDep As Single
    Dim unX1 As Single, unX2 As Single, unD1 As Single, unD2 As Single
    Dim unXecran As Single, unYecran As Single
    Dim unXecSuiv As Single, unYecSuiv As Single
    Dim unMinXecran As Single, unMaxXecran As Single
    Dim unMinYecran As Single, unMaxYecran As Single
    Dim uneDistMaxReelX As Single, uneDistMaxEcranX As Single
    Dim uneDistMaxReelY As Single, uneDistMaxEcranY As Single
    
    Screen.MousePointer = vbHourglass
    'Dépassement pour avant et après les limites de dessin
    'surtout lors des zooms cadré sur les repères début et fin de l'iti ref
    unDep = uneMargeB / 2
    
    'Détermination des min/max écran
    'Variables servant pour la conversion coordonnées réelles en écran
    uneDistMaxReelX = unMaxXreel - unMinXreel
    uneDistMaxEcranX = uneZoneDes.Width - uneMargeG - uneMargeD
    uneDistMaxReelY = unMaxYreel - unMinYreel
    uneDistMaxEcranY = uneZoneDes.Height - uneMargeH - uneMargeB

    'Conversion en coordonnées Y écran des distances réelles
    'Les Y sont orientés vers le bas, donc le max réel correspondant au max écran
    'est inférieur au min écran correspondant au min réel
    unMaxYecran = uneMargeH
    unMinYecran = uneZoneDes.Height - uneMargeB
    
    'Conversion en coordonnées X écran des temps ou vitesses réels
    unMinXecran = uneMargeG
    unMaxXecran = uneZoneDes.Width - uneMargeD
    
    'Dessin de la courbe DT des parcours utilisés
    For i = 1 To uneColPar.Count
        Set unParcours = uneColPar(i)
        If unParcours.monIsUtil Then
            'Si c'est le parcours choisi parmi une sélection multiple
            'le parcours est dessiné en trait épais en trait fin sinon
            If unIndParChoisi = i Then
                uneZoneDes.DrawWidth = TraitGros
            Else
                uneZoneDes.DrawWidth = TraitFin
            End If
            'Dessin de la courbe distance/temps du parcours
            uneCouleur = unParcours.maCouleur
            unNbPoints = unParcours.monNbPas
            'Récup des données pour une courbe temps/distance
            unD1 = 0
            unX1 = 0
            'Conversion des distances des décimètres au mètre
            unD2 = unParcours.monTabDist(1) / 10 * unParcours.monCoefEta
            'Conversion des temps des dixièmes de seconde au minute
            unX2 = unParcours.monFirstPas / 600
            
            'Conversion en coordonnées écrans des coordonnées réelles
            'du premier point =(unX1, unD1) et du deuxième point =(unX2, unD2)
            unXecran = ConvertirEnEcran(unMinXecran, unX1 - unMinXreel, uneDistMaxReelX, uneDistMaxEcranX)
            unYecran = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD1, uneDistMaxReelY, uneDistMaxEcranY)
            unXecSuiv = ConvertirEnEcran(unMinXecran, unX2 - unMinXreel, uneDistMaxReelX, uneDistMaxEcranX)
            unYecSuiv = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD2, uneDistMaxReelY, uneDistMaxEcranY)
            'Dessin de la courbe du premier segment
            's'il est entre le min et le max y écran
            'min y écran  > max y écran car les y écran orientés vers le bas en Y,
            'donc aprés conversion donnée réelle en écran le max devient < au min
            'If unYecran <= unMinYecran And unYecran >= unMaxYecran Then
            If (unYecran <= unMinYecran + unDep And unYecran >= unMaxYecran - unDep) Or (unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran) Then
                uneZoneDes.Line (unXecran, unYecran)-(unXecSuiv, unYecSuiv), uneCouleur
            End If
            'Stockage pour le segment suivant
            unXecran = unXecSuiv
            unYecran = unYecSuiv
            
            For j = 2 To unNbPoints - 1
                'Calcul du point suivant pour la courbe temps/distance
                'ou la courbe vitesse/distance
                'Cumul des distances et Conversion des distances
                'des décimètres au mètre
                unD2 = unD2 + unParcours.monTabDist(j) / 10 * unParcours.monCoefEta
                'Cumul des temps et Conversion du pas des secondes en minute
                unX2 = unX2 + unParcours.monPasMesure / 60
                
                'Conversion en coordonnées écrans des coordonnées réelles
                'du point suivant
                unXecSuiv = ConvertirEnEcran(unMinXecran, unX2 - unMinXreel, uneDistMaxReelX, uneDistMaxEcranX)
                unYecSuiv = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD2, uneDistMaxReelY, uneDistMaxEcranY)
                'Dessin de la courbe segment par segment
                's'il est entre le min et le max y écran
                'min y écran  > max y écran car les y écran orientés vers le bas en Y,
                'donc aprés conversion donnée réelle en écran le max devient < au min
                If (unYecran <= unMinYecran + unDep And unYecran >= unMaxYecran - unDep) Or (unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran) Then
                    uneZoneDes.Line (unXecran, unYecran)-(unXecSuiv, unYecSuiv), uneCouleur
                End If
                'Stockage pour le segment suivant
                unXecran = unXecSuiv
                unYecran = unYecSuiv
            Next j
            
            'Calcul du dernier point pour la courbe temps/distance
            'ou la courbe vitesse/distance
            'Conversion des distances des décimètres au mètre
            unD2 = unParcours.maDistPar / 10 * unParcours.monCoefEta
            'Conversion du pas des dixièmes de secondes en minute
            unX2 = unParcours.maDuree / 600
            'Conversion en coordonnées écrans des coordonnées réelles
            'du point suivant
            unXecSuiv = ConvertirEnEcran(unMinXecran, unX2 - unMinXreel, uneDistMaxReelX, uneDistMaxEcranX)
            unYecSuiv = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD2, uneDistMaxReelY, uneDistMaxEcranY)
            'Dessin de la courbe segment par segment
            's'il est entre le min et le max y écran
            'min y écran  > max y écran car les y écran orientés vers le bas en Y,
            'donc aprés conversion donnée réelle en écran le max devient < au min
            If (unYecran <= unMinYecran + unDep And unYecran >= unMaxYecran - unDep) Or (unYecSuiv <= unMinYecran And unYecSuiv >= unMaxYecran) Then
                uneZoneDes.Line (unXecran, unYecran)-(unXecSuiv, unYecSuiv), uneCouleur
            End If
        
            'Dessin des repères topés le long du parcours
            For j = 0 To UBound(unParcours.monTabAbsRep)
                If j = 0 Then
                    'Cas du repère de fin de mesure
                    'Conversion des distances des décimètres au mètre
                    unD1 = unParcours.maDistPar / 10 * unParcours.monCoefEta
                    'Conversion des temps des dixièmes de seconde au minute
                    unX1 = unParcours.maDuree / 600
                Else
                    'Cas des repères topés sur le parcours
                    'Conversion des distances des décimètres au mètre
                    unD1 = unParcours.monTabAbsRep(j) / 10 * unParcours.monCoefEta
                    'Conversion des temps des dixièmes de seconde au minute
                    unX1 = unParcours.monTabTempsRep(j) / 600
                End If
                
                'Conversion en coordonnées écrans des coordonnées réelles
                'du repère topé = (unX1, unD1)
                unXecran = ConvertirEnEcran(unMinXecran, unX1 - unMinXreel, uneDistMaxReelX, uneDistMaxEcranX)
                unYecran = ConvertirEnEcran(unMaxYecran, unMaxYreel - unD1, uneDistMaxReelY, uneDistMaxEcranY)
                'Dessin du repère topé, un carré centré sur (X1,D1), taille = TailleRep/2 x TailleRep/2
                's'il est entre le min et le max y écran
                'min y écran  > max y écran car les y écran orientés vers le bas en Y,
                'donc aprés conversion donnée réelle en écran le max devient < au min
                If (unYecran <= unMinYecran + unDep And unYecran >= unMaxYecran - unDep) Then
                    'uneZoneDes.Line (unXecran - TailleRep / 2, unYecran - TailleRep / 2)-(unXecran + TailleRep / 2, unYecran + TailleRep / 2), uneCouleur, BF
                    uneZoneDes.Line (unXecran - TailleRep * 0.75, unYecran - TailleRep * 0.75)-(unXecran + TailleRep * 0.75, unYecran + TailleRep * 0.75), uneCouleur, BF
                End If
            Next j
        End If
    Next i
    
    'Remise en trait fin des dessins
    uneZoneDes.DrawWidth = TraitFin
    
    Screen.MousePointer = vbDefault
End Sub

Public Function CouperParcoursEntreD1D2(unPar As Parcours, uneForm As Form, unNbRepIti As Integer, unIndRepItiDeb As Integer, unIndRepItiFin As Integer, unCoefEtirement As Single) As Boolean
    'Procédure modifiant les données d'un parcours en ne conservant que les
    'valeurs comprises entre deux abscisses curvilignes ou distances D1 et D2
    'Les valeurs sont mises à zéros en D1 et comptés jusqu'à D2 multiplié par
    'un coef d'étirement pour cadrer entre D1 et D2 (cf form frmImportMTB)
    
    '**************************************************************************
    'Les données sont stockées dans le parcours donnée par la variable globale
    'monParToImport et la valeur de retour est VRAI si la coupure a pu se faire
    'et FAUX sinon
    '**************************************************************************
    
    Dim uneD0 As Single, uneD1 As Single, uneD2 As Single
    Dim unNbRepTop As Integer, i As Long, unNbRepInDblTop As Integer
    Dim unNbRepTmp As Integer, unNbRepTotal As Integer
    Dim unNbDblTopDep As Integer, uneAbsCurv0 As Single
    Dim uneAbsCurv As Single, uneAbsCurvDeb As Single
    Dim uneAbsCurvPred As Single, uneDistPar As Long
    Dim unTempsDeb As Long, unTempsFin As Long
    Dim unPas As Long, unTemps As Long, unMappingRep As Boolean
    Dim unNumPasDeb As Long, unNumPasFin As Long
    Dim unRepTop1 As Integer, unRepTop2 As Integer
    Dim uneTolMax As Integer, unEcartMax As Integer
    Dim unTabAbsRep As Variant, unTabTmpRep As Variant
    Dim uneColAbs As New Collection, unIndColRepDeb As Integer
    Dim unNbRepCreer As Integer
    
    'Initialisation à VRAI du résultat de la coupure
    CouperParcoursEntreD1D2 = True
    
    uneD1 = CSng(Format(uneForm.ShapeRep(unIndRepItiDeb).Tag))
    uneD2 = CSng(Format(uneForm.ShapeRep(unIndRepItiFin).Tag))
    'On fait en sorte que D1 <= D2
    If uneD1 > uneD2 Then
        uneD0 = uneD1
        uneD1 = uneD2
        uneD2 = uneD0
    End If
    'Conversation des distances des mètres au décimètres car les données
    'des repères topés et des pas de mesures sont en décimètres
    uneD1 = uneD1 * 10
    uneD2 = uneD2 * 10
    
    'Calcul du coefficient d'étalonnage
    'On divise le coef d'étalonnage du parcours car c'est l'itinéraire
    'de réference que l'on a étiré dans frmImportMTB
    monParToImport.monCoefEta = unPar.monCoefEta / unCoefEtirement
    'Récupération des données communes
    monParToImport.monNom = unPar.monNom
    monParToImport.monIsUtil = unPar.monIsUtil
    monParToImport.maCouleur = unPar.maCouleur
    monParToImport.monEnqueteur = unPar.monEnqueteur
    monParToImport.monNumVeh = unPar.monNumVeh
    monParToImport.maMeteo = unPar.maMeteo
    monParToImport.maDate = unPar.maDate
    monParToImport.monJourSemaine = unPar.monJourSemaine
    monParToImport.monHeureDebut = unPar.monHeureDebut
    monParToImport.monPasMesure = unPar.monPasMesure
    monParToImport.monNumVeh = unPar.monNumVeh
    
    'Récupération de la tolérance max en pourcentage permettant de dire
    'si deux repères sont distincts (diff des abs curv / un des abs curv)
    'uneTolMax = mesOptions.maTolLong
    
    'Recherche des repères confondus à 25 m (= 250 decimètres) près pour la
    'correspondance entre les repères iti ref et ceux du parcours à importer
    uneTolMax = 250
    'et de l'écart max en décimètre indiquant si on a un double top ou pas
    unEcartMax = mesOptions.monEcartMax * 10
    'Mise à vrai de l'indication de la coïncidence des repères parcours
    'avec ceux de l'itinéraire de référence
    unMappingRep = True
    
    'Récupération du nombre de repères toppés entre D1 et D2 du parcours,
    'de leur abscisse curviligne et de leur temps de passage en comptant à
    'partir de D1
    unRepTop1 = 0
    uneAbsCurvPred = -1000
    uneAbsCurvDeb = -1000
    unNbRepInDblTop = 0 'Nombre de repère parcours faisant partie d'un double top
                        'souvent 2
    unRepTop2 = 0
    unNbRepTop = UBound(unPar.monTabAbsRep)
    'Compactage du nombre de repères itinéraire entre les repères début et fin
    'de l'itinéraire de référence car les repères ne sont pas classés par ordre
    'croissant ou décroissant
    unNbRepTmp = 0
    unNbRepTotal = 0
    unIndColRepDeb = 0
    For i = 1 To unNbRepIti
        'Conversion des mètres en décimètres
        uneD0 = CSng(Format(uneForm.ShapeRep(i).Tag) * 10)
        'Ajout dans la collection des abs curv de repères de l'iti ref
        'en ordonnant par ordre croissant
        For j = 1 To unNbRepTmp
            If uneD0 <= uneColAbs(j) Then
                uneColAbs.Add uneD0, , j
                Exit For
            End If
        Next j
        'Cas où aucun n'est plus petit ==> c'est le plus grand, mis en fin
        'Pour le premier rep iti inséré j vaut 1
        If j = unNbRepTmp + 1 Then uneColAbs.Add uneD0
        unNbRepTmp = unNbRepTmp + 1
        'Comptage du nb de repères entre D1 et D2 et stockage de l'indice
        'du repère début dans la collection des abs curv triées
        If uneD1 - unEcartMax <= uneD0 And uneD0 <= uneD2 + unEcartMax Then
            'Incrémentation du nombre de repères iti ref entre d1 et d2
            unNbRepTotal = unNbRepTotal + 1
            'Stockage de l'indice correspondant au repère début = abs min
            'on ne stocke que le premier trouvé
            If unIndColRepDeb = 0 Then
                If Abs(uneD0 - uneD1) <= uneTolMax Then
                    unIndColRepDeb = i
                End If
            End If
        End If
    Next i

    'Allocation dynamique des tableaux liés aux repères topés
    unTabAbsRep = monParToImport.monTabAbsRep
    ReDim unTabAbsRep(1 To unNbRepTop + 1)
    '+1 au cas où le début ne soit pas un des repères du parcours
    'à importer
    unTabTmpRep = monParToImport.monTabTempsRep
    ReDim unTabTmpRep(1 To unNbRepTop + 1)
    
    unNbRepTmp = 0
    unNbDblTopDep = 0 'Nombre de double top au départ pour les supprimer
    For i = 1 To unNbRepTop
        uneAbsCurv = unPar.monTabAbsRep(i) * unPar.monCoefEta
        uneAbsCurv0 = uneAbsCurvDeb * unPar.monCoefEta
        If unRepTop1 = 0 Or Abs(uneAbsCurv - uneAbsCurv0) <= unEcartMax Then
            'Le test Abs(uneAbsCurv - uneAbsCurv0) <= unEcartMax permet de trouver
            'les cas de double tops au départ pour ne pas les prendre en compte
            If Abs(uneAbsCurv - uneD1) <= uneTolMax Then
                'Cas où on trouve le premier repère du parcours importé qui
                'correspond au repère début de l'itinéraire de référence
                If unRepTop1 = 0 Then unRepTop1 = i
                'Détermination en décimètres de l'abs curv du début
                'et de son temps de passage
                uneAbsCurvDeb = unPar.monTabAbsRep(i)
                unTempsDeb = unPar.monTabTempsRep(i)
                'On incrémente le nombre de repère trouvé entre D1 et D2
                unNbRepTmp = unNbRepTmp + 1
            End If
        ElseIf unRepTop2 = 0 Then
            'On incrémente le nombre de repère trouvé entre D1 et D2
            unNbRepTmp = unNbRepTmp + 1
            If Abs(uneAbsCurv - uneD2) <= uneTolMax And unNbRepTmp = unNbRepTotal + unNbRepInDblTop Then
                'Cas où on trouve le premier repère du parcours importé qui
                'correspond au repère fin de l'itinéraire de référence et en ayant
                'trouvé autant de repère qu'entre D1 et D2 en tenant compte des
                'double tops, on évite les problèmes de précision
                unRepTop2 = i
                'Détermination en décimètres de l'abs curv de fin
                'et de son temps de passage
                unTempsFin = CLng(unPar.DonnerTempsPassage(uneAbsCurv)) 'uneD2))
            ElseIf unRepTop2 = 0 And uneAbsCurv > uneD2 Then
                'Cas du premier repère du parcours importé dépassant le début
                'et avec aucun repère coincidant avec la fin trouvé avant
                unRepTop2 = i
                'Détermination en décimètres de l'abs curv de fin
                'et de son temps de passage
                unTempsFin = CLng(unPar.DonnerTempsPassage(uneD2))
            End If
        End If
        If unRepTop1 > 0 And uneAbsCurv < uneD2 + uneTolMax Then
            'Remplissage du tableau des abscisses curvilignes
            'pour les repères ne dépassant pas le repère fin (< uneD2)
            'Récupération de l'abs curv du repère de l'itinéraire de référence
            'chargé censé coïncidé avec le repère du parcours à importer en
            'décompte le nombre de repère parcours qui sont des doubles tops
            uneD0 = uneColAbs(i - unRepTop1 - unNbRepInDblTop + unIndColRepDeb)
            'déjà en décimètres en tenant de l'indice du premier repère parcours
            'coincidant avec le repère début itinéraire et son indice dans la
            'collection des abs curv triées et des double tops
            
            'uneD0 = CSng(Format(uneForm.ShapeRep(i - unNbRepInDblTop - unRepTop1 + unIndRepItiDeb).Tag))
            'Conversion des mètres en décimètres
            'uneD0 = uneD0 * 10
            If Abs(uneAbsCurv - uneAbsCurvPred) <= uneTolMax Then 'unEcartMax Then
                'Cas d'un repère de parcours qui est un double top
                'même abs que le précédent
                If unRepTop1 + 1 + unNbDblTopDep = i Then
                    'Cas d'un double top au début
                    unNbDblTopDep = unNbDblTopDep + 1
                End If
                unTabAbsRep(i + 1 - unRepTop1 - unNbDblTopDep) = unPar.monTabAbsRep(i) - uneAbsCurvDeb
                unTabTmpRep(i + 1 - unRepTop1 - unNbDblTopDep) = unPar.monTabTempsRep(i) - unTempsDeb
                'Incrémentation du nbre de repères dans le dernier double top trouvé
                unNbRepInDblTop = unNbRepInDblTop + 1
            ElseIf Abs(uneAbsCurv - uneD0) <= uneTolMax Then
                'Cas où on trouve le repère du parcours importé qui
                'correspond au repère de l'itinéraire de référence
                unTabAbsRep(i + 1 - unRepTop1 - unNbDblTopDep) = unPar.monTabAbsRep(i) - uneAbsCurvDeb
                unTabTmpRep(i + 1 - unRepTop1 - unNbDblTopDep) = unPar.monTabTempsRep(i) - unTempsDeb
                'Stockage du nouveau abscisse précédent
                uneAbsCurvPred = uneAbsCurv
            Else
                'Mise à faux de l'indication de la coïncidence des repères parcours
                'avec ceux de l'itinéraire de référence
                unMappingRep = False
                Exit For
                'uneReponse = MsgBox("Le parcours à importer possède un ou plusieurs repères ne coïncidant pas avec ceux de l'itinéraire de référence chargé." + Chr(13) + Chr(13) + "Voulez-vous une correction automatique des repères du parcours à importer ?", vbOKCancel + vbCritical, "Erreur d'import")
                'If uneReponse = vbCancel Then
                    'Arrêt de la coupure, Mise à FAUX du résultat de la coupure
                    'CouperParcoursEntreD1D2 = False
                    'Exit Function
                'End If
            End If
        End If
        If unRepTop2 > 0 Then
            'Remplissage du dernier repère topé
            If uneAbsCurv > uneD2 Then
                unTabAbsRep(i + 1 - unRepTop1 - unNbDblTopDep) = uneD2 / unPar.monCoefEta - uneAbsCurvDeb
                unTabTmpRep(i + 1 - unRepTop1 - unNbDblTopDep) = unTempsFin - unTempsDeb
            End If
            'Sortie si on a trouvé ou dépassé le repère de fin
            Exit For
        End If
    Next i
    
    If unMappingRep = False Then
        uneReponse = MsgBox("Le parcours à importer possède un ou plusieurs repères ne coïncidant pas avec ceux de l'itinéraire de référence chargé." + Chr(13) + Chr(13) + "Voulez-vous une correction automatique des repères du parcours à importer ?", vbOKCancel + vbCritical, "Erreur d'import")
        If uneReponse = vbCancel Then
            'Suppression en mémoire de la collection des abs curv
            ViderCollection uneColAbs
            Set uneColAbs = Nothing
            'Arrêt de la coupure, Mise à FAUX du résultat de la coupure
            CouperParcoursEntreD1D2 = False
            Exit Function
        Else
            'MsgBox "Cette fonction n'est pas encore disponible.", vbInformation
            'CouperParcoursEntreD1D2 = False
            'Exit Function
            
            'Calcul des repères topés à partir des repères de l'itinéraire de référence
            'si demandé par l'utilisateur
            
            'On ne garde que la taille maximale
            If unNbRepTotal >= unNbRepTop Then
                unNbMax = unNbRepTotal
            Else
                unNbMax = unNbRepTop
            End If
            ReDim Preserve unTabAbsRep(1 To unNbMax)
            ReDim Preserve unTabTmpRep(1 To unNbMax)
            'Initialisation
            unNbRepCreer = 0
            i = unIndColRepDeb
            j = 1
            j0 = 1
            While i <= unNbRepTotal
                'Parcours de tous les repères de l'itinéraire de référence
                If Abs(unPar.monTabAbsRep(j) - uneColAbs(i)) <= uneTolMax Then
                    'Le repère topé correspond au repère de l'itinéraire
                    'de référence
                    'Incrémentation du nombre de repères topés créés
                    'et affection de son abs curv et de son temps de passage
                    unNbRepCreer = unNbRepCreer + 1
                    unTabAbsRep(unNbRepCreer) = unPar.monTabAbsRep(j) - uneAbsCurvDeb
                    unTabTmpRep(unNbRepCreer) = unPar.monTabTempsRep(j) - unTempsDeb
                    j0 = j
                    'Recherche et stockage des double tops
                    For k = j + 1 To unNbRepTop
                        If Abs(unPar.monTabAbsRep(k) - uneColAbs(i)) <= unEcartMax Then
                            'Le repère topé correspond au repère de l'itinéraire
                            'de référence mais de type double top
                            'Incrémentation du nombre de repères topés créés
                            'et affection de son abs curv et de son temps de passage
                            unNbRepCreer = unNbRepCreer + 1
                            unTabAbsRep(unNbRepCreer) = unPar.monTabAbsRep(k) - uneAbsCurvDeb
                            unTabTmpRep(unNbRepCreer) = unPar.monTabTempsRep(k) - unTempsDeb
                            j0 = k
                            'Rajout dans le tableau d'une colonne pour stocker le double top
                            ReDim Preserve unTabAbsRep(1 To unNbMax + 1)
                            ReDim Preserve unTabTmpRep(1 To unNbMax + 1)
                        Else
                            Exit For
                        End If
                    Next k
                    'Incrémentation pour faire le repère suivant
                    i = i + 1
                    j = j0
                Else
                    If j >= unNbRepTop Then
                        'Arrivée au dernier repère topé qui ne va pas non plus
                        '==> Cas d'un repère de référence non topé, on le recrée
                        'Incrémentation du nombre de repères topés créés
                        'et affection de son abs curv et de son temps de passage
                        'à partir
                        unNbRepCreer = unNbRepCreer + 1
                        unTabAbsRep(unNbRepCreer) = uneColAbs(i) - uneAbsCurvDeb
                        unTabTmpRep(unNbRepCreer) = CLng(unPar.DonnerTempsPassage(uneColAbs(i))) - unTempsDeb
                        'Incrémentation pour faire le repère iti ref suivant
                        i = i + 1
                        j = j0
                    Else
                        'Incrémentation pour faire le repère parcours suivant
                        j = j + 1
                    End If
                End If
            Wend
        End If
    End If
    
    If unRepTop1 = 0 Then
        MsgBox "Le parcours à importer n'a pas de repère coïncidant avec le repère début de l'itinéraire de référence chargé.", vbCritical, "Erreur d'import"
        'Suppression en mémoire de la collection des abs curv
        ViderCollection uneColAbs
        Set uneColAbs = Nothing
        'Mise à FAUX du résultat de la coupure
        CouperParcoursEntreD1D2 = False
        Exit Function
    End If
    
    'Si pas de fin trouvée on prend tout entre le premier top trouvé
    'ou le repère début créé éventuel
    'et le dernier repère topé, sinon jusqu'au repère trouvé
    If unRepTop2 = 0 Then
        unRepTop2 = unNbRepTop + 1 - unRepTop1
        'Détermination en décimètres de l'abs curv de fin
        'et de son temps de passage
        unTempsFin = CLng(unPar.DonnerTempsPassage(uneD2))
        'MsgBox "Le parcours à importer n'a pas de repère coïncidant avec le repère fin de l'itinéraire de référence chargé.", vbCritical, "Erreur d'import"
        'Mise à FAUX du résultat de la coupure
        'CouperParcoursEntreD1D2 = False
        'Exit Function
    Else
        unRepTop2 = unRepTop2 + 1 - unRepTop1
    End If
    'Prise en compte des double tops trouvés au début
    unRepTop2 = unRepTop2 - unNbDblTopDep
    
    'Modification suite à l'insertion automatique des repères iti ref manquants
    If unMappingRep = False Then unRepTop2 = unNbRepCreer
    'Suppression en mémoire de la collection des abs curv
    ViderCollection uneColAbs
    Set uneColAbs = Nothing
    
    'On ne garde que les valeurs situées entre le début et la fin
    ReDim Preserve unTabAbsRep(1 To unRepTop2)
    ReDim Preserve unTabTmpRep(1 To unRepTop2)
    'Affectation des pointeurs sur le tableau
    'des abscisses curvilignes et des temps de passage des repères du parcours
    monParToImport.monTabAbsRep = unTabAbsRep
    monParToImport.monTabTempsRep = unTabTmpRep
    
    'Récupération des données par pas de mesure du parcours à importer
    unPar.DonnerInterDistance unTempsDeb, unNumPasDeb
    unPar.DonnerInterDistance unTempsFin, unNumPasFin
    If unNumPasFin = unNumPasDeb Then
        MsgBox "Le parcours à importer est à découper entre deux repères début et fin trop proches.", vbCritical, "Erreur d'import"
        'Mise à FAUX du résultat de la coupure
        CouperParcoursEntreD1D2 = False
        Exit Function
    Else
        'Calcul du nombre de pas de mesure
        monParToImport.monNbPas = unNumPasFin - unNumPasDeb + 1
    End If
    
    'Calcul du premier pas de mesure pour cela :
    '   Calcul du pas contenant le repère début
    '   Calcul du temps de passage à la fin du pas contenant le repère début
    If unNumPasDeb = 1 Then
        unPas = unPar.monFirstPas
    Else
        unPas = unPar.monPasMesure * 10
        'Conversion du pas de mesure des secondes en dixième de secondes
    End If
    unTemps = unPar.monFirstPas + unPas * (unNumPasDeb - 1)
    'Calcul du premier pas de mesure
    If unPas = 0 Then
        monParToImport.monTabDist(1) = unPar.monTabDist(unNumPasDeb)
    Else
        monParToImport.monTabDist(1) = (unTemps - unTempsDeb) / unPas * unPar.monTabDist(unNumPasDeb)
    End If
    'Calcul de la durée du dernier pas de mesure
    monParToImport.monFirstPas = unTemps - unTempsDeb
    
    'Alimentation des distances parcourues par pas autre que le premier
    'et le dernier et calcul de la distance parcourue
    uneDistPar = monParToImport.monTabDist(1)
    For i = unNumPasDeb + 1 To unNumPasFin - 1
        monParToImport.monTabDist(i + 1 - unNumPasDeb) = unPar.monTabDist(i)
        uneDistPar = uneDistPar + unPar.monTabDist(i)
    Next i
    
    'Calcul du premier pas de mesure pour cela :
    '   Calcul du temps de passage juste avant le pas contenant le repère fin
    '   Conversion du pas de mesure des secondes en dixième de secondes
    unTemps = unPar.monFirstPas + unPar.monPasMesure * 10 * (unNumPasFin - 2)
    '   Calcul du pas contenant le repère fin
    If unNumPasFin = unPar.monNbPas Then
        unPas = unPar.monLastPas
    Else
        unPas = unPar.monPasMesure * 10
        'Conversion du pas de mesure des secondes en dixième de secondes
    End If
    'Calcul du dernier pas de mesure
    If unPas = 0 Then
        monParToImport.monTabDist(monParToImport.monNbPas) = unPar.monTabDist(unNumPasFin)
    Else
        monParToImport.monTabDist(monParToImport.monNbPas) = CLng((unTempsFin - unTemps) / unPas * unPar.monTabDist(unNumPasFin))
    End If
    'fin du calcul de la distance parcourue = longueur du parcours
    monParToImport.maDistPar = uneDistPar + monParToImport.monTabDist(monParToImport.monNbPas)
    'Calcul de la durée
    monParToImport.maDuree = unTempsFin - unTempsDeb
    'Calcul de la durée du dernier pas de mesure
    monParToImport.monLastPas = unTempsFin - unTemps
End Function
