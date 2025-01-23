Attribute VB_Name = "M�dulo2"
' Created By: Breno Henrique
' Date: 01/2025
' Objective: Rotine para verifica��o da vers�o da imagem para garantir a atualiza��o constante das imagens

Sub VersionVerified()
    ' Rotina para verifica��o da vers�o das imagens
    
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
    Set wsSheet = wbBook.Worksheets("C�lculos NBR9062 2017")

    ' Temp
    Dim result() As String
    Dim lastElement As String
    Dim versao As Integer
    
    versao = wsSheet.Cells(26, 14).Value

    ' Verificar se h� algo selecionado
    If wdApp.Selection.Type = wdSelectionInlineShape Then
        ' Se a sele��o for uma InlineShape (imagem inline)
        Set selectedInlineShape = wdApp.Selection.InlineShapes(1)
        
        result = Split(selectedInlineShape.AlternativeText, ":")
        lastElement = result(UBound(result))
        
        ' Verified version
        If CInt(lastElement) = versao Then
            MsgBox "A Imagem est� atualizada"
            
        Else
            MsgBox "A Imgem est� desatualizadas: versao no documento " & lastElement & " enquanto na planilha est� na vers�o " & versao
        End If
               
    Else
        MsgBox "Nenhuma imagem selecionada ou objeto incompat�vel."
    End If
    
End Sub
