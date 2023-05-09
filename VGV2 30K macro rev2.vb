
Sub STATISTIQUE_FICHIER_MTU_30K_VGV2_PTH_PANOPLIE_LISTING()
'
' ENREGISTREMENT Macro
'
' DEMANDE DE SELECTION D'UN FICHIER A OUVRIR POUR L'ENREGISTREMENT

Dim Fichier As Variant
Dim Fichier2 As String
Dim Fichier3 As String
Dim Fichier4 As String
Dim Chemin1 As String
Dim Chemin2 As String
Dim MasterChemin As String
Dim Dossier As String
Dim Excel As String
Dim fs As Object
Dim Wb As Workbook
Dim new_ChaineNUM As String
Dim NUM As String
Dim dt As Date
Dim Dat As String
Dim Chaine As String
Dim LenChaine As Integer
Dim new_Chaine As String
Dim PosOF As Integer
Dim PosNum As Integer
Dim OF As String
Dim ChaineF2 As String
Dim LenChaineF2 As Integer
Dim PosRécap As Integer
Dim LenChaineR As Integer
Dim new_ChaineR As String
Dim LenChaineRecap As Integer
Dim ChaineR As String
Dim ChaineF3 As String
Dim PosPAL As Integer
Dim new_ChainePAL As String
Dim LenChainePAL As Integer
Dim PAL As String
Dim TRIDIM As String
Dim Derog As Boolean
Dim Fine As Boolean
Dim Epais As Boolean
Dim PtTabSupIntraP As Boolean
Dim PtTabSupIntraN As Boolean
Dim PtTabSupExtraP As Boolean
Dim PtTabSupExtraN As Boolean
Dim PtTabInfIntraP As Boolean
Dim PtTabInfIntraN As Boolean
Dim PtTabInfExtraP As Boolean
Dim PtTabInfExtraN As Boolean
Dim CordeG As Boolean
Dim CordeP As Boolean
Dim CordeS As Boolean
Dim Rayon As Boolean
Dim Dforme As Boolean
Dim Fut As Boolean
Dim Ctrl As Boolean
Dim NbDefaut As Integer
Dim ChaineF4 As String
Dim PosREP As Integer
Dim new_ChaineREP As String
Dim LenChaineREP As Integer
Dim REP As String
Dim BoEvent As Boolean, BoSaut As Boolean
Dim Wbmaster As String
Dim Cible As Object, Valeur As Object

' On conserve d'abord les configurations existantes
  BoEvent = Application.EnableEvents
  BoSaut = ActiveSheet.DisplayPageBreaks
  Wbmaster = ActiveWorkbook.Name
  MasterChemin = ActiveWorkbook.Path
  MasterChemin = Replace(MasterChemin, "\Excel", "")
' Affiche la boîte de dialogue "Sélectionner un fichier" en . XLSX
  ChDrive "G:"
  ChDir MasterChemin
  Fichier = Application.GetOpenFilename("Tous les fichiers (*.xlsx),*.*")
' On sort si aucun fichier n'a été sélectionné ou si l'utilisateur
' a cliqué sur le bouton "Annuler", ou sur la croix de fermeture.
  If Fichier = False Then Exit Sub
' Affiche le chemin et le nom du fichier sélectionné.
' On force les configurations
  Application.ScreenUpdating = False
  Application.DisplayStatusBar = False
  Application.EnableEvents = False
  ActiveSheet.DisplayPageBreaks = False
' Récupération du chemin des fichiers
' Ouverture du fichier sélectionné
  Workbooks.Open Filename:=Fichier
' Définit le chemin du répertoire contenant les fichiers
  Chemin1 = ActiveWorkbook.Path
  Set Cible = CreateObject("Scripting.FileSystemObject")
  Set Valeur = Cible.GetFile(ActiveWorkbook.FullName)
' Récupération de la date de création du fichier
    dt = Valeur.DateCreated
' Définit le répertoire Excel de copie des fichiers
  Dossier = "Excel"
' Vérifie si le dossier Excel existe
  If Len(Dir(Dossier, vbDirectory)) > 0 Then
  DossierExiste = True
  Else
  DossierExiste = False
' MsgBox "Le dossier n'existe pas..."
' Création du dossier Excel pour déplacement les fichiers traités
  Set fs = CreateObject("Scripting.FileSystemObject")                     ' initialisation de la variable
  If fs.FolderExists(Dossier) Then                                        ' Si le repertoire existe donc rien a faire
  Else: fs.CreateFolder Dossier                                           ' Sinon le repertoire n'existe pas donc on le créer
  End If
  Set fs = Nothing
  End If
' Boucle sur tous les fichiers xls du répertoire.
' Chemin des répertoires de travail
  Chemin1 = Chemin1 & "\"
  Chemin2 = Chemin1 & Dossier & "\"
' Fermeture du fichier sélectionné
  ActiveWorkbook.Close False
' Boucle sur le répertoire chemin1 et ses fichiers Excel en .XLSX
  Fichier2 = Dir(Chemin1 & "*.xlsx")
  Do While Len(Fichier2) > 0
  Debug.Print Chemin1 & Fichier2
  Debug.Print Chemin2 & Fichier2
  PosNum = InStr(14, Fichier2, "_")
' Ouverture du premier fichier trouvé
  Set Wb = Workbooks.Open(Chemin1 & Fichier2)
' Récupère le chemin du fichier
  Chaine = ActiveWorkbook.Path
  Fichier4 = Left(Fichier2, PosNum + 4)
  Fichier4 = Fichier4 & ".xlsx"
' Recherche le mot OF dans la chaine de caracrtère et donne ça position
  PosOF = InStr(Chaine, "OF")
' Détermine la longueur de la chaine de caractères - 1
  LenChaine = Len(Chaine) - 1
' nouvelle chaine = ancienne chaine - position mot OF à droite
  new_Chaine = Right(Chaine, LenChaine - PosOF)
' Sélectionne les 6 caractères de gauche de la nouvelle chaine
  OF = Left(new_Chaine, 6)
' Chaine de carractères complète avec fichier
  ChaineF2 = (Chaine & Fichier4)
' Sélectionne les caractères de droite de la nouvelle chaine
  ChaineR = Left(Right(ChaineF2, 8), 3)
' Sélectionne les 6 caractères de gauche de la nouvelle chaine            27
  NUM = Left(ChaineR, 8)
' Chaine de carractères complète avec fichier
  ChaineF3 = (Chaine & Fichier4)
' Sélectionne les 2 caractères de gauche à partir des 12 derniers caractères du fichier (numéro de palette)
  PAL = Left(Right(ChaineF3, 12), 2)
' Supprimer des _ dans le numéro de palette
  PAL = Replace(PAL, "_", "")
' Recherche le mot rep dans la chaine de caracrtère
' Chaine de carractères complète avec fichier
  ChaineF4 = (Chaine)
' et donne ça position
  PosREP = InStr(ChaineF4, "rep")
' Sélectionne les 6 caractères de gauche de la nouvelle chaine
  REP = Right(Left(ChaineF4, 90), 4)
' Copier coller la colonne FF des résultats du contrôle ( Ecarts ) et numéro de tridim
    TRIDIM = Range("A151")
    Columns("F:F").Select
    Selection.Copy
' Coller les résultats dans la colonne TT
    Windows(Wbmaster).Activate
    Columns("T:T").Select
    Selection.Insert Shift:=xlToRight
' COLLER LE NUMERO DE L'OF EN BOUT DE COLONNE T124
    Range("T501") = OF
' COLLER LE NUMERO DE TRIDIM EN COLONNE T497
    TRIDIM = Replace(TRIDIM, "936", "Crysta V1")
    TRIDIM = Replace(TRIDIM, "974", "Crysta V2")
    TRIDIM = Replace(TRIDIM, "277", "Crysta S")
    Range("T497") = TRIDIM
' COLLER LE NUMERO DE PIECE EN BOUT DE COLONNE T357
    Range("T500") = NUM
' Colle la date du fichier dans la celule T126 et réalise la mise en forme
    Range("T503") = dt
    Range("T503").NumberFormat = "dd/mm/yy;@"
' COLLER LE NUMERO DE LA PALETTE EN BOUT DE COLONNE T502
    Range("T502") = REP
' COLLER LE NUMERO DE REPERE du LOT de PIECE EN BOUT DE COLONNE T504
    Range("T504") = PAL
' Supprime le contenu de la cellule ( gain de temps sur la macro! )
    Range("B10000").Select
    Selection.Delete Shift:=xlUp
    Range("B1").Select
' Fermeture du fichier
    Wb.Close False
' Copie du Fichier dans un autre répertoire
    Dat = Replace(dt, ":", "-")
    Dat = Replace(Dat, "/", "-")
    Fichier4 = Replace(Fichier4, ".xlsx", " ")
    Fichier3 = Fichier4 & Dat & ".xlsx"
    FileCopy Chemin1 & Fichier2, Chemin2 & Fichier3
    'Détruit le fichier du répertoire après la copie pour ne plus le sélectionner
    Kill Chemin1 & Fichier2
        Fichier2 = Dir()
' Fin de boucle sur Fichier2
' Active et copie les cellules de contrôle de la colonne d'après, et les copies dans la colonne TT
    Range("U520:U650").Select
    Selection.Copy
    Range("T520").Select
    ActiveSheet.Paste
    ' Contrôle extensions tolérances individuelles
    'Remise à 0 des defauts
    Derog = False
    Fine = False
    Epais = False
    PtTabSupIntraP = False
    PtTabSupIntraN = False
    PtTabSupExtraP = False
    PtTabSupExtraN = False
    PtTabInfIntraP = False
    PtTabInfIntraN = False
    PtTabInfExtraP = False
    PtTabInfExtraN = False
    CordeG = False
    CordeP = False
    CordeS = False
	Rayon = False
    Dforme = False
    Fut = False
    Ctrl = False
    ' Si déro
    For Each Cell In Range("T115:T122")
    If (Cell.Value >= 0.175 And Cell.Value <= 0.2) Or (Cell.Value <= -0.175 And Cell.Value >= -0.2) Then
    Derog = True
    End If
    Next
 
 'Si pieces fines
 For Each Cell In Range("T3:T7,T10:T14")
 If Cell.Value < -0.12 Then
 Fine = True
 End If
 If Cell.Value > 0.14 Then
 Epais = True
 End If
 Next
 
 'Si Points de tablette
 For Each Cell In Range("T115:T116")
 If Cell.Value > 0.2 Then
 PtTabSupIntraP = True
 End If
 If Cell.Value < -0.2 Then
 PtTabSupIntraN = True
 End If
 Next
 For Each Cell In Range("T117:T118")
 If Cell.Value > 0.2 Then
 PtTabSupExtraP = True
 End If
 If Cell.Value < -0.2 Then
 PtTabSupExtraN = True
 End If
 Next
 For Each Cell In Range("T119:T120")
 If Cell.Value > 0.2 Then
 PtTabInfIntraP = True
 End If
 If Cell.Value < -0.2 Then
 PtTabInfIntraN = True
 End If
 Next
 For Each Cell In Range("T121:T122")
 If Cell.Value > 0.2 Then
 PtTabInfExtraP = True
 End If
 If Cell.Value < -0.2 Then
 PtTabInfExtraN = True
 End If
 Next
 
 'Si corde trop grande
 For Each Cell In Range("T38:T42")
 If Cell.Value > 0.331 Then
 CordeG = True
 End If
 'Si corde trop petite
 If Cell.Value < -0.331 Then
 CordeP = True
 End If
 Next
 'Si somme des cordes hors tolérence
 If Range("T604").DisplayFormat.Interior.Color = RGB(255, 0, 0) Then
 CordeS = True
 End If
 'Si defaut rayon d'attache
 For Each Cell In Range("T137:T140")
 If Cell.Value > 0.25 Or Cell.Value < -0.25 Then
 Rayon = True
 End If
 Next
 
 'Si défaut de formes
 For Each Cell In Range("T17:T21")
 If Cell.Value < -0.51 Or Cell.Value > 0.51 Then
 Dforme = True
 End If
 Next
 For Each Cell In Range("T24:T28")
 If Cell.Value < -0.405 Or Cell.Value > 0.405 Then
 Dforme = True
 End If
 Next
 For Each Cell In Range("T31:T35")
 If Cell.Value < -0.8 Or Cell.Value > 0.8 Then
 Dforme = True
 End If
 Next
 For Each Cell In Range("T45:T49,T52:T56,T59:T63,T66:T70")
 If Cell.Value < -0.12 Or Cell.Value > 0.12 Then
 Dforme = True
 End If
 Next
 For Each Cell In Range("T73,T77,T80,T84,T94,T98,T101,T105")
 If Cell.Value < -0.075 Or Cell.Value > 0.075 Then
 Dforme = True
 End If
 Next
 For Each Cell In Range("T74:T76,T81:T83,T95:T97,T102:T104")
 If Cell.Value < -0.06 Or Cell.Value > 0.06 Then
 Dforme = True
 End If
 Next
 For Each Cell In Range("T87,T91,T108,T112")
 If Cell.Value < 0 Or Cell.Value > 0.15 Then
 Dforme = True
 End If
 Next
 For Each Cell In Range("T88:T90,T109:T111")
 If Cell.Value < 0 Or Cell.Value > 0.12 Then
 Dforme = True
 End If
 Next
 
 'Si defaut de fût
 For Each Cell In Range("T125,T126,T143:T145,T148:T150")
 If Cell.Value > 1.2 Or Cell.Value < -1.2 Then
 Fut = True
 End If
 Next
 
 'Si valeur aberrante
 For Each Cell In Range("T3:T112")
 If Cell.Value > 2 Or Cell.Value < -2 Then
 Ctrl = True
 End If
 Next
 
 'Conditions pour déterminer comment est la pièce
 NbDefaut = CInt(Derog) + CInt(Fine) + CInt(Epais) + CInt(PtTabSupIntraP) + CInt(PtTabSupIntraN) + CInt(PtTabSupExtraP) + CInt(PtTabSupExtraN) + CInt(PtTabInfIntraP) + CInt(PtTabInfIntraN) + CInt(PtTabInfExtraP) + CInt(PtTabInfExtraN) + CInt(CordeG) + CInt(CordeP) + CInt(Rayon) + CInt(Dforme) + CInt(Fut) + CInt(Ctrl) + Cint(CordeS)

 If Ctrl = True Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(242, 242, 242)
 Range("T498").Interior.Color = RGB(255, 0, 255)
 Range("T499") = "ctrl x2"
 GoTo FinProc
 End If
 
 If NbDefaut = -1 Then
 If Derog = True And Range("T618").DisplayFormat.Interior.Color <> RGB(255, 0, 0) And Range("T632").DisplayFormat.Interior.Color <> RGB(255, 0, 0) Then
 Range("T498").Interior.Color = RGB(0, 176, 240)
 Range("T499") = "Dero"
 End If
 If Derog = True And Range("T618").DisplayFormat.Interior.Color <> RGB(255, 0, 0) And Range("T632").DisplayFormat.Interior.Color = RGB(255, 0, 0) Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 44
 Range("T499") = "Pt Tab"
 End If
 If Fut = True Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 44
 Range("T499") = "fut"
 End If
 If Fine = True Or CordeG = True Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 44
 Range("T499") = "TriboA"
 End If
 If Epais = True Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 44
 Range("T499") = "épaisse"
 End If
 If PtTabSupIntraP = True Or PtTabSupIntraN = True Or PtTabSupExtraP = True Or PtTabSupExtraN = True Or PtTabInfIntraP = True Or PtTabInfIntraN = True Or PtTabInfExtraP = True Or PtTabInfExtraN = True Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 44
 Range("T499") = "Pt Tab"
 End If
 If CordeP = True Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 3
 Range("T499") = "REBUT"
 End If
 If Rayon = True Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 44
 Range("T499") = "rayon"
 End If
 If Dforme = True Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 3
 Range("T499") = "df forme"
 End If
 GoTo FinProc
 End If
 
 If NbDefaut < -1 Then
 If Epais = True Or CordeP = True Or Rayon = True Or Dforme = True Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 3
 Range("T499") = "REBUT"
 GoTo FinProc
 End If
 For Each Cell In Range("T119:T122")
 If Cell.Value > 0.15 And Fut = True And Range("T125") < 0 And Range("T126") < 0 Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 3
 Range("T499") = "REBUT"
 GoTo FinProc
 End If
 Next
 If (Fut = True And PtTabSupIntraP = True) Or (Fut = True And PtTabSupExtraP = True) Or (Fut = True And PtTabInfIntraP = True) Or (Fut = True And PtTabInfExtraP = True) Or (Fut = True And Derog = True) Or (PtTabSupIntraP = True And PtTabSupExtraP = True) Or (PtTabSupIntraN = True And PtTabSupExtraN = True) Or (PtTabInfIntraP = True And PtTabInfExtraP = True) Or (PtTabInfIntraN = True And PtTabInfExtraN = True) Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 3
 Range("T499") = "REBUT"
 GoTo FinProc
 End If
 If CordeG = True Or Fine = True Or CordeS = True Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 44
 Range("T499") = "TriboA"
 GoTo FinProc
 End If
 If PtTabSupIntraP = True Or PtTabSupIntraN = True Or PtTabSupExtraP = True Or PtTabSupExtraN = True Or PtTabInfIntraP = True Or PtTabInfIntraN = True Or PtTabInfExtraP = True Or PtTabInfExtraN = True Then
 Range("T498") = 1
 Range("T1:T504").Interior.Color = RGB(217, 217, 217)
 Range("T498").Interior.ColorIndex = 44
 Range("T499") = "Pt Tab"
 GoTo FinProc
 End If
 End If
 
 If NbDefaut = 0 Then
 Range("T498").Interior.Color = RGB(0, 176, 80)
 Range("T499") = "ok"
 End If
 
FinProc:
 Range("T497:T504").HorizontalAlignment = xlCenter
 Range("T497:T504").Font.Bold = True
 
 Loop

' Vous avez votre code qui est defini ici avant d'arriver a la fin les configurations sont restaurees
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = BoEvent
ActiveSheet.DisplayPageBreaks = BoSaut

End Sub
