Attribute VB_Name = "CopyControl"
'-----------------------------------------------------------------
'
'  CCMiTemps.BAS
' Programme gestion de protection CopyControl pour MiTemps
' pouvant servir de mod�le pour d'autres produits
'
' Origine : CERTU - Mars 98
' Mise � jour : CETE de l'Ouest (A.VIGNAUD) - Avril 98
'               CERTU (F.TRIFILETTI) - Juillet 99
'-----------------------------------------------------------------

Option Explicit

'Variable disant si on travaille en version prot�g�e ou non
Public maProtectVersion As Boolean

' Variable globale pour A_Propos
Public SerialNumber As String

' Variable globale � placer dans le module principal
'Public VersionDemo As Boolean

' A param�trer selon l'application
Private Const BonCode = "its02Mit3.0"
Private Const BonneDLL = "MITEMPS.DLL" 'Mettre le nom de dll en majuscule
'En majuscule car on va comparer cette variable au nom du produit
'prot�g� par CopyControl or ce nom est mis d'office en majuscule
Private Const NomAppli = "MiTemps"
Private LongueurCode As Integer
Private LongueurNomDLL As Integer

Private Const MinTime = 15 ' une v�rif toutes les 15mn

' Conserve l'heure et le jour de la derni�re modification pour MyCCToujours()
Private MyDerModif As Integer
Private MyNewDay As Integer

' Disque porteur de la protection
Private disquePorteur As Integer

' M�morisation du disque et du r�pertoire courants
Private saveDrive As String
Private saveDir As String  '(utile uniqt ds l'envt de d�veloppement)

Const Long256 As Long = 256

'D�claration de la structure
Type CCMB
  B1 As String * 1
  B2 As String * 1
  B3 As String * 1
  B4 As String * 1
  Func As String * 1
  Rcodelo As String * 1
  Rcodehi As String * 1
'  Rcode As Integer ' ADAPTATION 16bits pour faciliter Rcode = -1 (Integer = 2 octets)
  Drive As String * 1
  Dir As String * 4
  Vers1 As String * 1
  Vers2 As String * 1
  SN As String * 2
'  SN As Integer ' � la place de String * 2; pour faciliter la lecture (mais attention au +/-)
  Pcode As String * 9
  Pname As String * 13
  CCSN As String * 2
  Master As String * 1
  DrType As String * 1
  Copies1 As String * 1
  Copies2 As String * 1
  InitCopies As String * 2
  Useslo As String * 1
  Useshi As String * 1
  IUseslo As String * 1
  IUseshi As String * 1
  ExpD As String * 1
  ExpM As String * 1
  ExpYlo As String * 1
  ExpYhi As String * 1
  NotreDecalage As String * 4        '4 octets pris sur Remainder
  MsgSecurit As String * 256         '256 octets pris sur Remainder
  
  Remainder As String * 198   '458 - 260 // taille de la structure = 512o'
End Type

'D�claration des variables'
Private myCC As CCMB

'D�claration de la DLL en 32 bits d'apr�s <exemple>
Declare Function ccdll Lib "Mitemps.dll" Alias "CC32" (CC As CCMB) As Integer

Private Sub ClearStruct(lpCC As CCMB)
'Initialisation de la structure d'apr�s doc
' ... et obligatoirement remettre Dir � NULL pour A: en cas de modification sur C:)
  lpCC.B1 = "C"
  lpCC.B2 = "C"
  lpCC.B3 = "M"
  lpCC.B4 = "B"
  lpCC.Func = Chr$(0)
'  lpCC.Rcode = 0
  lpCC.Rcodelo = Chr$(255) ' l'ensemble des 2 octets (FF)
  lpCC.Rcodehi = Chr$(255) ' donne -1 (utile pour la version 16 bits)
  lpCC.Drive = Chr$(0)
  lpCC.Dir = String$(4, 0)
  lpCC.Vers1 = Chr$(0)
  lpCC.Vers2 = Chr$(0)
'  lpCC.SN = 0
  lpCC.SN = String$(2, 0)
  lpCC.Pcode = String$(9, 0)
  lpCC.Pname = String$(13, 0)
  lpCC.CCSN = String$(2, 0)
  lpCC.Master = Chr$(0)
  lpCC.DrType = Chr$(0)
  lpCC.Copies1 = Chr$(0)
  lpCC.Copies2 = Chr$(0)
  lpCC.InitCopies = String$(2, 0)
  lpCC.Useslo = Chr$(0)
  lpCC.Useshi = Chr$(0)
  lpCC.IUseslo = Chr$(0)
  lpCC.IUseshi = Chr$(0)
  lpCC.ExpD = Chr$(0)
  lpCC.ExpM = Chr$(0)
  lpCC.ExpYlo = Chr$(0)
  lpCC.ExpYhi = Chr$(0)
  lpCC.NotreDecalage = String$(4, 0)
  lpCC.MsgSecurit = String$(256, 0)
  lpCC.Remainder = String$(198, 0)
End Sub


'utilisation de la protection'
Public Function ProtectCheck(uneFonc As Integer) As Integer
    ' Appel de appelDLL appelant la DLL de CopyControl
    'pour v�rification de la protection
    
    ' uneFonc  = 0 - v�rification + inscription en tant qu'utilisateur
    '            2 - v�rification seule
    
    ' disque   = 0 - disque courant : recherche sur CC_PATH, chemin programme (o� se trouve la DLL),
    '                                 R�p. de travail, R�p racine
    '                                 sinon cf myCC.Dir
    '            1 - disquette A
    '
    'Retourne le code d'erreur de v�rif de la dll prot�g�
    'si ce code = 0 ==> Protection valide, sinon protection non valide
    
    Dim comp_status As Integer
    Dim comp_statusA As Integer
    Dim CChaine As String
    Dim NotreCode As String
    Dim flag As Boolean
    Static unNbAppelVerifProt As Integer
    
    'Les Lignes suivantes utile uniquement en d�veloppement sous VB
    'car la sortie de l'exe lanc� de VB ne rend pas le jeton r�seau
    'On retourne comme si la protection �tait pr�sente
    'Si on veut tester avec protection if faut changer la valeur
    'maProtectVersion dans la fonction Main d'un des .bas
    'MsgBox "Appel de protectCheck avec " + Format(uneFonc)
    If maProtectVersion = False Then
        'Cas d'une version non prot�g�e
        ProtectCheck = 0
        If unNbAppelVerifProt = 0 Then
            MsgBox "Version non prot�g�e de " + App.Title + " pour test interne au CERTU uniquement. NE PAS LA DIFFUSER", vbInformation
            unNbAppelVerifProt = 1
        End If
        Exit Function
    End If
    
    'Initialisation des codes et dll � v�rifier
    LongueurCode = Len(BonCode)
    LongueurNomDLL = Len(BonneDLL)
    
    'on commence par regarder sur le disque courant
    ChangeRep   ' on rend courant le disque de l'application
    disquePorteur = 0
    comp_status = appelDLL(uneFonc, disquePorteur, flag)
    RetrouveRep
    If flag Then ProtectCheck = -20000: Exit Function ' DLL non trouv�e
    
    'MsgBox "1er appel " & CStr(comp_status)
    disquePorteur = Asc(Left(App.Path, 1)) - 64   ' utilis� par les futurs appels de CCToujours
    
    If (comp_status = -28) Then   'correspond au Msg : Transf�rez le jeton dans le r�pertoire courant !
    'on regarde la disquette A
        disquePorteur = 1
        comp_statusA = appelDLL(uneFonc, disquePorteur, flag)
    '    MsgBox "2� appel " & CStr(comp_statusA)
        Select Case comp_statusA
                Case 0
                    comp_status = 0          'on valide
                Case -57                     'erreur apparaissant si prot�t�g� en �criture
                    comp_status = -5700
                Case -67, -26
                    comp_status = comp_statusA
        End Select
    End If

    '-------- R�cup�ration du Num�ro de S�rie
    '         ... et Test du nom de la DLL original, et du Code
    If (comp_status = 0) Then   'myCC est soit celle du disque courant, soit celle de la disquette
        SerialNumber = Str$(Asc(myCC.SN) + Long256 * Asc(Mid(myCC.SN, 2)))
        NotreCode = Mid(myCC.MsgSecurit, 2, LongueurCode)
        'ATTENTION : � adapter au message de s�curit�
        '       BonCode : its00Str1.0
        '       Left  : nombre de caract�res de la partie visible et cach� + 1 (pour le caract�re null ~Z)
        '       Right : nombre de caract�res de la partie cach�
        If (Left$(myCC.Pname, LongueurNomDLL) <> BonneDLL Or NotreCode <> BonCode) Then comp_status = -19000
    End If
        
    'Retour du code de v�rif protection
    ProtectCheck = comp_status
    
    '------- Messages d'erreur et sortie
    If (comp_status <> 0) Then
        Select Case comp_status
            Case -19
                CChaine = "Produit non install� !"
            Case -26
                CChaine = "Le num�ro de licence ne correspond pas"
            Case -28
                CChaine = "Jeton introuvable"
            Case -35
                CChaine = "V�rification impossible : le disque est prot�g� en �criture !"
            Case -5700
                CChaine = "V�rification impossible : la disquette est prot�g�e en �criture !"
            Case -67
                CChaine = "Veuillez recommencer plus tard" & Chr(13) & "Trop d'utilisateurs sont pr�sents !"
            Case Else
                CChaine = "Erreur n� " & Str$(comp_status) & Chr(13) & NomAppli & " n'a pas trouv� la protection"
        End Select
        
        
        MsgBox CChaine, vbCritical, NomAppli + " : Gestion de la Protection"
        Exit Function
    End If
    
    MyDerModif = 60 * Hour(Now) + Minute(Now)
    MyNewDay = Day(Now)
    
End Function


Private Sub ChangeRep()
' Repositionnement �ventuel du disque courant sur celui de l'application
' pour �tre s�r de trouver la protection

  If Mid(CurDir, 2, 1) = ":" Then
    If Left(CurDir, 1) <> Left(App.Path, 1) Then
      saveDrive = Left(CurDir, 1)
      ChDrive Left(App.Path, 1)
    Else
      saveDrive = ""
    End If
  End If
 
' Exit Sub
 
' Repositionnement �ventuel du r�pertoire courant sur celui de l'application
' pour �tre s�r de trouver la DLL
' utile uniquement en environnement de d�veloppement
  If StrComp(CurDir, App.Path, 1) <> 0 Then
    saveDir = CurDir
    ChDir App.Path
  Else
    saveDir = ""
  End If

End Sub

Private Sub RetrouveRep()

  If saveDrive <> "" Then ChDrive saveDrive
  If saveDir <> "" Then ChDir saveDir

End Sub

Private Function appelDLL(ByVal fonction As Integer, ByVal disque As Integer, Absent As Boolean) As Integer
    ' Appel de la DLL de CopyControl pour v�rification de la protection
    
    ' fonction = 0 - v�rification + inscription en tant qu'utilisateur
    '            2 - v�rification seule
    
    ' disque   = 0 - disque courant : recherche sur CC_PATH, chemin programme (o� se trouve la DLL),
    '                                 R�p. de travail, R�p racine
    '                                 sinon cf myCC.Dir
    '            1 - disquette A
    
    Dim souris As Integer   'sauvegarde de la forme souris
    Dim Msg As String
    
    souris = Screen.MousePointer
    Screen.MousePointer = vbHourglass    ' sablier
    
    ClearStruct myCC       ' r�initialisation de la structure
    myCC.Func = Chr$(fonction)
    myCC.Drive = Chr$(disque)
    
    On Error GoTo GestErr
    appelDLL = ccdll(myCC)   'appel de la DLL
    Screen.MousePointer = souris
    
    Exit Function

'----gestion de l'absence de la DLL
GestErr:
  Msg = Err.Description
  MsgBox Msg, vbCritical, NomAppli + " : Gestion de la protection"
  Absent = True
  Resume Next

End Function

