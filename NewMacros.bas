Attribute VB_Name = "NewMacros"


Sub Pdfbox()
UserForm1.Show
End Sub



Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
ActiveDocument.ExportAsFixedFormat ActiveDocument.Path & "\PDF\1. Coperta.pdf", _
        wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportFromTo, 1, 1, wdExportDocumentWithMarkup, _
        False, False, wdExportCreateHeadingBookmarks, True, False, False
ActiveDocument.ExportAsFixedFormat ActiveDocument.Path & "\PDF\2. Borderou.pdf", _
        wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportFromTo, 2, 2, wdExportDocumentWithMarkup, _
        False, False, wdExportCreateHeadingBookmarks, True, False, False
ActiveDocument.ExportAsFixedFormat ActiveDocument.Path & "\PDF\1. Coperta.pdf", _
        wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportFromTo, 3, 3, wdExportDocumentWithMarkup, _
        False, False, wdExportCreateHeadingBookmarks, True, False, False
ActiveDocument.ExportAsFixedFormat ActiveDocument.Path & "\PDF\1. Coperta.pdf", _
        wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportFromTo, 4, 4, wdExportDocumentWithMarkup, _
        False, False, wdExportCreateHeadingBookmarks, True, False, False
ActiveDocument.ExportAsFixedFormat ActiveDocument.Path & "\PDF\1. Coperta.pdf", _
        wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportFromTo, 5, 5, wdExportDocumentWithMarkup, _
        False, False, wdExportCreateHeadingBookmarks, True, False, False

End Sub
Sub LawFormat()
Attribute LawFormat.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
'
'

    Selection.WholeStory
    Selection.Font.Name = "Arial"
    Selection.Font.Size = 10
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = -654262273
    End With
    With Selection.Find
        .Text = "^013Articolul ([0-9]@)^013"
        .Replacement.Text = "Art. \1^013"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = -654262273
    End With
    With Selection.Find
        .Text = "^013Articolul ([0-9]@^0094[0-9]@)^013"
        .Replacement.Text = "Art. \1^013"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = -721371137
    End With
    With Selection.Find
        .Text = "^013(\([0-9]@\))"
        .Replacement.Text = "^013\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = -721371137
    End With
    With Selection.Find
        .Text = "^013(\([0-9]@?[0-9]@\))"
        .Replacement.Text = "^013\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = -687816705
    End With
    With Selection.Find
        .Text = "^013([a-z]@\))"
        .Replacement.Text = "^013\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = -687816705
    End With
    With Selection.Find
        .Text = "^13([a-z]@^0094[0-9]\))"
        .Replacement.Text = "^13\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013\((la)?@^013"
        .Replacement.Text = "^013"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^013Not" & ChrW(259) & "^013\*\)?@^013"
        .Replacement.Text = "^013"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^013^013"
        .Replacement.Text = "^013"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013(Art\. [0-9]@)^013"
        .Replacement.Text = " ^013\1 - "
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013(Art\. [0-9]@^0094[0-9]@)^013"
        .Replacement.Text = " ^013\1 - "
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013"
        .Replacement.Text = " ^013^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
      Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t(Capitolul)"
        .Replacement.Text = "^013^t\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^t(Titlul)"
        .Replacement.Text = "^013^t\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorRed
    With Selection.Find
        .Text = "(Abrogat)"
        .Replacement.Text = "\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = True
        .Color = -738148353
    End With
    With Selection.Find
        .Text = "^t(Titlul?@)^13"
        .Replacement.Text = "^t\1^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^t(Capitolul?@)^13"
        .Replacement.Text = "^t\1^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub PrintPDFRangeMailing()
Attribute PrintPDFRangeMailing.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
'
'
strTotal = InputBox("Numarul total de pagini: ")
strFrom = InputBox("De la pagina: ")
strTo = InputBox("La pagina: ")
For i = 1 To strTotal
ActiveDocument.MailMerge.DataSource.ActiveRecord = i
strSeIdentificaCu = ActiveDocument.MailMerge.DataSource.DataFields("Se_identifica_cu").Value
strNumePtCerere = ActiveDocument.MailMerge.DataSource.DataFields("NumePtCerere").Value
ActiveDocument.ExportAsFixedFormat ActiveDocument.Path & "\PDF\" & strSeIdentificaCu & " - " & strNumePtCerere, _
        wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportFromTo, strFrom, strTo, wdExportDocumentWithMarkup, _
        False, False, wdExportCreateHeadingBookmarks, True, False, False
Next i
    
End Sub

