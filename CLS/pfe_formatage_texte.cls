VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'================================================================
' MACRO TEXTE - Mise en forme du texte PFE (norme APA 7)
' Styles Normal, Titres 1-6, numerotation romaine, nettoyage, TOC
' (la mise en forme des tableaux se trouve dans le fichier separe
'  pfe_formatage_tableaux.cls)
'================================================================

Sub FormaterTextePFE()

    Dim doc As Document
    Set doc = ActiveDocument

    Application.ScreenUpdating = False

    AppliquerStylesTexte doc          ' Normal, Titres, TOC
    AppliquerNumerotationTitre1 doc   ' Numerotation en chiffres romains (I, II, III...) pour Titre 1
    NettoyerDocument doc              ' Espaces multiples, cesure, compat
    MettreAJourTOC doc                ' Actualise la table des matieres

    Application.ScreenUpdating = True

    MsgBox "Mise en forme du texte PFE terminee : Titre 1 numerote en chiffres romains," & vbCrLf & _
           "espace avant chaque niveau de titre, interligne 1.5 dans le texte.", vbInformation

End Sub


'================================================================
' 1. STYLES DE TEXTE (Normal, Titres 1-5, TOC 1-5)
'================================================================

Sub AppliquerStylesTexte(doc As Document)

    '--- Style Normal ---
    With doc.Styles("Normal")
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = False
        .Font.Italic = False

        With .ParagraphFormat
            .Alignment = wdAlignParagraphJustify
            .LineSpacingRule = wdLineSpace1pt5
            .SpaceBefore = 0
            .SpaceAfter = 6           ' espace entre les paragraphes normaux (hors tableaux)
            .FirstLineIndent = CentimetersToPoints(1.27)
        End With
    End With

    ' Securite : forcer l'espacement "SpaceAfter" sur tous les paragraphes
    ' utilisant le style Normal, SAUF ceux qui se trouvent dans un tableau
    ' (les cellules de tableau sont gerees dans le fichier des tableaux).
    Dim paraNormal As Paragraph
    For Each paraNormal In doc.Paragraphs
        If paraNormal.Style = doc.Styles("Normal") Then
            If paraNormal.Range.Information(wdWithInTable) = False Then
                paraNormal.Format.SpaceAfter = 6
            End If
        End If
    Next paraNormal

    '--- Titres APA 7 (5 niveaux officiels + niveau 6) ---
    ' Chaque niveau recoit desormais un espace AVANT le titre, degressif
    ' du niveau 1 (le plus grand espace) au niveau 6 (le plus petit).
    ' Niveau 1 : centre, gras, TAILLE AUGMENTEE (plus gros que le texte) + numerote en chiffres romains
    ConfigurerTitre doc, "Heading 1", 14, True, False, wdAlignParagraphCenter, 0, False, 24
    ' Niveau 2 : aligne a gauche, gras
    ConfigurerTitre doc, "Heading 2", 13, True, False, wdAlignParagraphLeft, 0, False, 18
    ' Niveau 3 : aligne a gauche, gras italique
    ConfigurerTitre doc, "Heading 3", 12, True, True, wdAlignParagraphLeft, 0, False, 14
    ' Niveau 4 : en retrait, gras, texte suit sur la meme ligne (terminer le titre par un point)
    ConfigurerTitre doc, "Heading 4", 12, True, False, wdAlignParagraphLeft, 1.27, True, 12
    ' Niveau 5 : en retrait, gras italique, texte suit sur la meme ligne (terminer le titre par un point)
    ConfigurerTitre doc, "Heading 5", 12, True, True, wdAlignParagraphLeft, 1.27, True, 12
    ' Niveau 6 : en retrait, gras italique, texte suit sur la meme ligne (comme niveau 5)
    ConfigurerTitre doc, "Heading 6", 12, True, True, wdAlignParagraphLeft, 1.27, True, 10

    '--- Table des matieres (styles TOC 1-5) ---
    ConfigurerTOC doc, "TOC 1", 0
    ConfigurerTOC doc, "TOC 2", 0.5
    ConfigurerTOC doc, "TOC 3", 1
    ConfigurerTOC doc, "TOC 4", 1.5
    ConfigurerTOC doc, "TOC 5", 2

End Sub


Sub ConfigurerTitre(doc As Document, _
                     nomStyle As String, _
                     taille As Integer, _
                     gras As Boolean, _
                     italique As Boolean, _
                     align As WdParagraphAlignment, _
                     retrait As Double, _
                     texteSuite As Boolean, _
                     Optional espaceAvant As Single = 0)

    Dim st As Style
    Set st = RecupererStyle(doc, nomStyle)

    With st
        .Font.Name = "Times New Roman"
        .Font.Size = taille
        .Font.Bold = gras
        .Font.Italic = italique

        With .ParagraphFormat
            .Alignment = align
            .LeftIndent = CentimetersToPoints(retrait)
            .FirstLineIndent = 0   ' pas de retrait suspendu : le titre + le texte qui suit s'alignent ensemble

            .SpaceBefore = espaceAvant   ' espace au-dessus du titre (rien apres : le texte doit suivre directement)
            .SpaceAfter = 0
            .LineSpacingRule = wdLineSpace1pt5
            .KeepWithNext = True
            .OutlineLevel = NiveauPlanDepuisNomStyle(nomStyle)  ' place correctement le titre dans le plan / la TOC
        End With
    End With

    ' Niveaux 4 et 5 (titres "en ligne") : le texte du corps suit immediatement
    ' le titre sur la meme ligne, apres un point. Word ne peut pas ajouter ce
    ' point automatiquement (il ne connait pas le texte que tu vas taper) :
    ' pense a terminer toi-meme le titre par un "." avant d'ecrire la suite.

    ' Securite : si des paragraphes utilisant deja ce style ont un espace
    ' avant impose manuellement (formatage direct), on l'ecrase pour que
    ' l'espacement du style s'applique vraiment partout (uniquement hors tableaux).
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        If para.Style = st Then
            If para.Range.Information(wdWithInTable) = False Then
                para.Format.SpaceBefore = espaceAvant
                para.Format.SpaceAfter = 0
            End If
        End If
    Next para

End Sub


Function NiveauPlanDepuisNomStyle(nomStyle As String) As WdOutlineLevel

    Select Case nomStyle
        Case "Heading 1": NiveauPlanDepuisNomStyle = wdOutlineLevel1
        Case "Heading 2": NiveauPlanDepuisNomStyle = wdOutlineLevel2
        Case "Heading 3": NiveauPlanDepuisNomStyle = wdOutlineLevel3
        Case "Heading 4": NiveauPlanDepuisNomStyle = wdOutlineLevel4
        Case "Heading 5": NiveauPlanDepuisNomStyle = wdOutlineLevel5
        Case "Heading 6": NiveauPlanDepuisNomStyle = wdOutlineLevel6
        Case Else: NiveauPlanDepuisNomStyle = wdOutlineLevelBodyText
    End Select

End Function


Sub ConfigurerTOC(doc As Document, nomStyle As String, retrait As Double)

    Dim st As Style
    Set st = RecupererStyle(doc, nomStyle)

    With st
        .Font.Name = "Times New Roman"
        .Font.Size = 12

        With .ParagraphFormat
            .LeftIndent = CentimetersToPoints(retrait)
            .FirstLineIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = wdLineSpace1pt5
        End With
    End With

End Sub


Function RecupererStyle(doc As Document, nomStyle As String) As Style

    On Error Resume Next
    Set RecupererStyle = doc.Styles(nomStyle)

    If RecupererStyle Is Nothing Then
        Set RecupererStyle = doc.Styles.Add(Name:=nomStyle, Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0

End Function


'================================================================
' 1bis. NUMEROTATION EN CHIFFRES ROMAINS POUR TITRE 1 (I., II., III. ...)
'================================================================

Sub AppliquerNumerotationTitre1(doc As Document)

    Dim lt As ListTemplate
    Dim para As Paragraph
    Dim premierTitre As Paragraph

    Set lt = ListGalleries(wdOutlineNumberGallery).ListTemplates(1)

    With lt.ListLevels(1)
        .NumberFormat = "%1."                      ' ex : "I.", "II.", "III."
        .NumberStyle = wdListNumberStyleUppercaseRoman
        .NumberPosition = CentimetersToPoints(0)
        .TextPosition = CentimetersToPoints(1.5)
        .TabPosition = CentimetersToPoints(1.5)
        .Alignment = wdListLevelAlignLeft
        .TrailingCharacter = wdTrailingTab
        .LinkedStyle = "Heading 1"                  ' lie la numerotation au style Titre 1
        .StartAt = 1
    End With

    ' On accroche la liste au premier paragraphe en Titre 1 : grace au
    ' LinkedStyle ci-dessus, TOUS les paragraphes en Titre 1 du document
    ' (presents et futurs) seront numerotes automatiquement en chiffres
    ' romains, dans l'ordre, sans avoir a le refaire a chaque titre.
    For Each para In doc.Paragraphs
        If para.Style = doc.Styles("Heading 1") Then
            Set premierTitre = para
            Exit For
        End If
    Next para

    If Not premierTitre Is Nothing Then
        premierTitre.Range.ListFormat.ApplyListTemplateWithLevel _
            ListTemplate:=lt, _
            ContinuePreviousList:=False, _
            ApplyTo:=wdListApplyToWholeList, _
            DefaultListBehavior:=wdWord10ListBehavior
    End If

End Sub


'================================================================
' 2. NETTOYAGE DU DOCUMENT (espaces, cesure, compatibilite)
'================================================================

Sub NettoyerDocument(doc As Document)

    ' Supprimer les espaces multiples dans tout le document
    With doc.Content.Find
        .ClearFormatting
        .Text = "  "
        .Replacement.ClearFormatting
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
    End With

    ' Desactiver la cesure automatique
    doc.AutoHyphenation = False

    ' Desactiver l'option de compatibilite qui exagere les espaces
    On Error Resume Next
    doc.Compatibility(wdWordPerfectJustification) = False
    On Error GoTo 0

End Sub


'================================================================
' 3. MISE A JOUR DE LA TABLE DES MATIERES
'================================================================

Sub MettreAJourTOC(doc As Document)

    Dim toc As TableOfContents

    For Each toc In doc.TablesOfContents
        toc.Update
    Next toc

End Sub
