Attribute VB_Name = "Módulo1"
Sub exportaPDF()

    Dim minhaData As String
    minhaData = Format(Date, "dd") & " de" & Application.Text(Date, "[$-pt-BR] mmmm") & " de " & Format(Date, "yyyy")
    Range("A2").Value = minhaData
    

    Set objWord = CreateObject("Word.Application")
    
    objWord.Visible = True
    
    Set arqModelo = objWord.Documents.Open(ThisWorkbook.Path & "\modeloWord.docx")
    Set conteudoDoc = arqModelo.Application.Selection
    
    For i = 1 To 6
        conteudoDoc.Find.Text = Cells(1, i).Value
        conteudoDoc.Find.Replacement.Text = Cells(2, i).Value
        conteudoDoc.Find.Execute Replace:=wdReplaceAll
    Next
    
    arqModelo.SaveAs2 (ThisWorkbook.Path & "\Documentos Gerados" & "\" & Cells(2, 2).Value & ".docx")
    arqModelo.Close
    objWord.Quit
    
    Set objWord = Nothing
    Set arqModelo = Nothing
    Set conteudoDoc = Nothing
    
    MsgBox ("Documento Gerado com Sucesso!")
    
End Sub

