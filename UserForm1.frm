VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "PDFBox"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4245
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

If CheckBox1.Value = True Then
ActiveDocument.ExportAsFixedFormat ActiveDocument.Path & "\PDF\" & TextBox_a.Value & ".pdf", _
        wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportFromTo, c1.Value, c2.Value, wdExportDocumentWithMarkup, _
        False, False, wdExportCreateHeadingBookmarks, True, False, False
End If
        
If CheckBox2.Value = True Then
ActiveDocument.ExportAsFixedFormat ActiveDocument.Path & "\PDF\" & TextBox_c.Value & ".pdf", _
        wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportFromTo, b1.Value, b2.Value, wdExportDocumentWithMarkup, _
        False, False, wdExportCreateHeadingBookmarks, True, False, False
End If


If CheckBox3.Value = True Then
ActiveDocument.ExportAsFixedFormat ActiveDocument.Path & "\PDF\" & TextBox_j.Value & ".pdf", _
        wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportFromTo, m1.Value, m2.Value, wdExportDocumentWithMarkup, _
        False, False, wdExportCreateHeadingBookmarks, True, False, False
End If


If CheckBox4.Value = True Then
ActiveDocument.ExportAsFixedFormat ActiveDocument.Path & "\PDF\" & TextBox1.Value & ".pdf", _
        wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportFromTo, TextBox2.Value, TextBox3.Value, wdExportDocumentWithMarkup, _
        False, False, wdExportCreateHeadingBookmarks, True, False, False
End If
        
End Sub

Private Sub Label1_Click()

End Sub

Private Sub CommandButton2_Click()
CheckBox1.Value = False
CheckBox2.Value = False
CheckBox3.Value = False
CheckBox4.Value = False
End Sub

Private Sub UserForm_Click()

End Sub
