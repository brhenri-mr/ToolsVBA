Attribute VB_Name = "Módulo2"
' Created By: Breno Henrique
' Date: 01/2025
' Objective: Rotine para verificação da versão da imagem para garantir a atualização constante das imagens

Sub VersionVerified()
    ' Rotina para verificação da versão das imagens
    
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    Dim selectedInlineShape As InlineShape
    
    ' Referenciar o aplicativo e o documento Word ativo
    Set wdApp = Word.Application
    Set wdDoc = wdApp.ActiveDocument
    
    ' Excel obj
    Dim wsSheet As Worksheet
    Dim wbBook As Workbook
    Set wbBook = ThisWorkbook
    Set wsSheet = wbBook.Worksheets("Cálculos NBR9062 2017")

    ' Temp
    Dim result() As String
    Dim lastElement As String
    Dim versao As Integer
    
    versao = wsSheet.Cells(26, 14).Value

    ' Verificar se há algo selecionado
    If wdApp.Selection.Type = wdSelectionInlineShape Then
        ' Se a seleção for uma InlineShape (imagem inline)
        Set selectedInlineShape = wdApp.Selection.InlineShapes(1)
        
        result = Split(selectedInlineShape.AlternativeText, ":")
        lastElement = result(UBound(result))
        
        ' Verified version
        If CInt(lastElement) = versao Then
            MsgBox "A Imagem está atualizada"
            
        Else
            MsgBox "A Imgem está desatualizadas: versao no documento " & lastElement & " enquanto na planilha está na versão " & versao
        End If
               
    Else
        MsgBox "Nenhuma imagem selecionada ou objeto incompatível."
    End If
    
End Sub
