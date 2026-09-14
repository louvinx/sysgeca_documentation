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
' MACRO TABLEAUX - Mise en forme des tableaux PFE (norme APA 7)
' Tableau simple (quadrillage), alignement gauche (jamais justifie),
' interligne simple, AUCUN retrait, legendes centrees.
' (la mise en forme du texte/titres se trouve dans le fichier separe
'  pfe_formatage_texte.cls)
'================================================================

Sub FormaterTableauxPFE()

    Dim doc As Document
    Set doc = ActiveDocument

    Application.ScreenUpdating = False

    FormaterTousLesTableaux doc       ' Style + alignement gauche (jamais justifie) + interligne simple
    FormaterLegendesTableaux doc      ' Nom/legende des tableaux : centree, petite taille, italique
    SupprimerEspaceAvantTableaux doc  ' Aucun espace entre le texte/la legende et le tableau

    Application.ScreenUpdating = True

    MsgBox "Mise en forme des tableaux PFE terminee : tableau simple," & vbCrLf & _
           "alignement gauche, interligne simple, aucun retrait," & vbCrLf & _
           "legendes de tableaux centrees et reduites.", vbInformation

End Sub


'================================================================
' 1. MISE EN FORME DES TABLEAUX (jamais justifies)
'================================================================

Sub FormaterTousLesTableaux(doc As Document)

    Dim tbl As Table
    Dim cel As Cell
    Dim rw As Row
    Dim para As Paragraph
    Dim i As Long, j As Long
    Dim txt As String

    For Each tbl In doc.Tables

        '--- Style general : tableau simple (quadrillage) ---
        tbl.Style = "Table Grid"
        tbl.AllowAutoFit = True
        tbl.AutoFitBehavior wdAutoFitWindow
        tbl.Rows.Alignment = wdAlignRowCenter

        tbl.TopPadding = 4
        tbl.BottomPadding = 4
        tbl.LeftPadding = 5
        tbl.RightPadding = 5

        '--- Bordures simples : quadrillage complet, fines lignes noires ---
        With tbl.Borders
            .Enable = True
            .InsideLineStyle = wdLineStyleSingle
            .OutsideLineStyle = wdLineStyleSingle
            .InsideLineWidth = wdLineWidth050pt
            .OutsideLineWidth = wdLineWidth050pt
            .InsideColor = wdColorBlack
            .OutsideColor = wdColorBlack
        End With

        '--- Toutes les cellules : gauche, jamais justifie, AUCUN retrait,
        '    AUCUN espace avant/apres, interligne simple ---
        For Each cel In tbl.Range.Cells

            txt = cel.Range.Text
            txt = Replace(txt, Chr(13), "")
            txt = Replace(txt, Chr(7), "")

            If Trim(txt) = "" Then
                cel.Range.Text = " "
            End If

            With cel.Range
                .Font.Name = "Times New Roman"
                .Font.Size = 11
                .Font.Bold = False
                .Font.Color = wdColorBlack

                With .ParagraphFormat
                    .Alignment = wdAlignParagraphLeft
                    .SpaceBefore = 0   ' jamais d'espace entre paragraphes dans un tableau
                    .SpaceAfter = 0    ' jamais d'espace entre paragraphes dans un tableau
                    .LineSpacingRule = wdLineSpaceSingle  ' interligne simple
                    .FirstLineIndent = 0   ' aucun retrait
                    .LeftIndent = 0        ' aucun retrait
                End With
            End With

            cel.VerticalAlignment = wdCellAlignVerticalCenter

        Next cel

        '--- Premiere ligne = en-tete (sans fond de couleur, juste gras + ligne de separation) ---
        If tbl.Rows.Count > 0 Then
            tbl.Rows(1).HeadingFormat = True

            For j = 1 To tbl.Rows(1).Cells.Count
                With tbl.Rows(1).Cells(j)
                    .Shading.BackgroundPatternColor = RGB(230, 230, 230)
                    .Shading.Texture = wdTextureNone

                    With .Range
                        .Font.Name = "Times New Roman"
                        .Font.Size = 11
                        .Font.Bold = True
                        .Font.Color = wdColorBlack
                        With .ParagraphFormat
                            .Alignment = wdAlignParagraphCenter
                            .FirstLineIndent = 0
                            .LeftIndent = 0
                        End With
                    End With
                End With
            Next j
        End If

        '--- Corps du tableau : fond blanc uniforme (style academique, pas de bandes colorees) ---
        For i = 2 To tbl.Rows.Count
            For j = 1 To tbl.Rows(i).Cells.Count
                tbl.Rows(i).Cells(j).Shading.BackgroundPatternColor = wdColorWhite
                tbl.Rows(i).Cells(j).Shading.Texture = wdTextureNone
            Next j
        Next i

        '--- Hauteur minimale des lignes ---
        For Each rw In tbl.Rows
            rw.HeightRule = wdRowHeightAtLeast
            rw.Height = CentimetersToPoints(0.55)
        Next rw

        '--- Securite : forcer alignement gauche si un paragraphe est reste justifie ---
        For Each para In tbl.Range.Paragraphs
            If para.Alignment = wdAlignParagraphJustify Then
                para.Alignment = wdAlignParagraphLeft
            End If
        Next para

    Next tbl

End Sub


'================================================================
' 2. LEGENDES / NOMS DES TABLEAUX (style Caption / Legende)
'    -> CENTREE, petite taille de police, italique, interligne
'       simple. "Tableau X" au-dessus du tableau, collee au
'       tableau (pas d'espace apres la legende).
'================================================================

Sub FormaterLegendesTableaux(doc As Document)

    Dim st As Style
    Dim para As Paragraph

    ' Le style de legende s'appelle "Caption" en anglais et peut
    ' apparaitre sous le nom "Légende" dans un Word en francais.
    Set st = RecupererStyleLegende(doc)
    If st Is Nothing Then Exit Sub

    With st
        .Font.Name = "Times New Roman"
        .Font.Size = 10          ' plus petite que le texte courant (12)
        .Font.Bold = False
        .Font.Italic = True

        With .ParagraphFormat
            .Alignment = wdAlignParagraphCenter   ' legende centree
            .FirstLineIndent = 0
            .LeftIndent = 0
            .RightIndent = 0
            .LineSpacingRule = wdLineSpaceSingle
            .SpaceBefore = 6      ' petit espace au-dessus (separe du texte precedent)
            .SpaceAfter = 0       ' colle au tableau qui suit (aucun espace en haut du tableau)
            .KeepWithNext = True  ' la legende ne se separe jamais du tableau
            .OutlineLevel = wdOutlineLevelBodyText
        End With
    End With

    ' Securite : si certains paragraphes utilisent deja ce style mais
    ' avaient un retrait, une justification ou une taille appliques
    ' manuellement, on force le format directement dessus.
    For Each para In doc.Paragraphs
        If para.Style = st Then
            With para.Format
                .Alignment = wdAlignParagraphCenter
                .FirstLineIndent = 0
                .LeftIndent = 0
                .RightIndent = 0
            End With
            With para.Range.Font
                .Italic = True
                .Size = 10
            End With
        End If
    Next para

End Sub


Function RecupererStyleLegende(doc As Document) As Style

    On Error Resume Next
    Set RecupererStyleLegende = doc.Styles("Caption")
    If RecupererStyleLegende Is Nothing Then
        Set RecupererStyleLegende = doc.Styles("Légende")
    End If
    On Error GoTo 0

End Function


'================================================================
' 3. AUCUN ESPACE ENTRE LE TEXTE/LA LEGENDE ET LE TABLEAU
'================================================================

Sub SupprimerEspaceAvantTableaux(doc As Document)

    Dim tbl As Table
    Dim rngAvant As Range
    Dim derniereLigne As Long

    For Each tbl In doc.Tables

        If tbl.Range.Start > 0 Then
            Set rngAvant = doc.Range(0, tbl.Range.Start)

            If rngAvant.Paragraphs.Count > 0 Then
                derniereLigne = rngAvant.Paragraphs.Count
                With rngAvant.Paragraphs(derniereLigne).Range.ParagraphFormat
                    .SpaceAfter = 0
                End With
            End If
        End If

    Next tbl

End Sub
