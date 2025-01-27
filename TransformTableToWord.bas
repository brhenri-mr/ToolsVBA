Attribute VB_Name = "M�dulo1"
' Created By: Breno Henrique
' Date: 01/2025

'Objective:
' Interation between Excel tables and Word. The script allow the use to
' import the table as imagem That it fix perfect with the avaible space on the word


' Problems
' If the script does not working, may the VBA need to be configuret to that
' Ferramentas -> Prefer�ncias -> MicrosoftWordTools-> Check

' Backups can be founded at: https://github.com/brhenri-mr/ToolsVBA

Sub Export_Table_Word()
    
    Dim planName As String
    
    'Word objects.
    Dim wdbmRange As Word.Range
    Dim wdDoc As Word.Document
    Dim stWordReport As String
    
    'Excel objects.
    Dim wbBook As Workbook
    Dim wsSheet As Worksheet
    Dim rnReport As Range
    Dim piece As Range
    
    Dim firstRow As Long
    Dim maxRow As Integer
    Dim deltaP As Integer ' Posi��o variavel

    'Initialize the Excel objects.
    Set wbBook = ThisWorkbook
    Set wsSheet = wbBook.ActiveSheet
    planName = Cells(27, 14).Value & "CALCULOS"
    Set rnReport = wsSheet.Range(planName)

    
    'User Pattern
    Dim mark As String
    Dim choose As String
    Dim Path As String
    Dim rowSlice As Integer
    Dim versao As Integer
    Dim userName As String
    
    mark = Cells(23, 14).Value ' Name of the Bookmark on the word
    choose = Cells(24, 14).Value ' Basic a Boolean var to choose the document
    stWordReport = Cells(25, 14).Value 'Name of the existing Word doc.
    versao = Cells(26, 14).Value ' Sheets version
    userName = Cells(22, 14).Value


    firstRow = rnReport.row
    deltaP = firstRow
    
    maxRow = rnReport.Rows(1).row + rnReport.Rows.Count - 1
    

    ' Init Document's instance
    Set wdDoc = getWordDocument(choose, stWordReport, mark)
    
    ' Verified if the Bookmarks were register on the document
    If wdDoc.Bookmarks.Exists(mark) = True Then
        Set wdbmRange = wdDoc.Bookmarks(mark).Range
'--------------------------------------------------------------------------------------------
        While rowSlice < maxRow
        
            
            If deltaP = firstRow Then
                rowSlice = CalculateRowsPerSlice(rnReport, wdDoc, wdbmRange, True) + deltaP
                
            Else
                rowSlice = CalculateRowsPerSlice(rnReport, wdDoc, wdbmRange, False) + deltaP
                
            End If
            
            If rowSlice > maxRow Then
                rowSlice = maxRow + 1
            End If
            
            
            ' Piece of table
            Set piece = wsSheet.Range(wsSheet.Cells(deltaP, 1), wsSheet.Cells(rowSlice - 1, rnReport.Columns.Count))
            
            'Copy the report to the clipboard.
            piece.Copy
            
            'Select the range defined by the "Report" bookmark and paste in the report from clipboard.
            With wdbmRange
                .PasteSpecial Link:=False, _
                              DataType:=wdPasteMetafilePicture, _
                              Placement:=wdInLine, _
                              DisplayAsIcon:=False
            End With
            
            ' Insert Metadate on the img
            Call signIgm(wdDoc, versao, userName)
            
            ' Atualizando o loop
            Call MoveCursor(wdbmRange)
            wdDoc.Bookmarks.Add Name:=mark, Range:=wdbmRange
            deltaP = rowSlice ' Ele pega a nova posicao inicial
            
        Wend
'--------------------------------------------------------------------------------------------
        wdDoc.Bookmarks(mark).Delete
        Call finalizeWordDocument(choose, wdDoc, wdbmRange)

    Else
        MsgBox "Marcador n�o encontrado. Tem certeza que cadastrou um?"
    End If

End Sub

Function CalculateRowsPerSlice(excelRange As Range, wdDoc As Object, wdbmRange As Object, first As Boolean) As Integer
    ' Funcao para transformacao de linha no excel para unidades de medida no word

    Dim pageHeight As Single
    Dim markerPosition As Single
    Dim availableSpace As Single
    Dim internalSpace As Single
    
    Dim rowHeightExcel As Single
    Dim totalRows As Long
       
    ' Relative position on the page
    markerPosition = wdbmRange.Information(wdVerticalPositionRelativeToPage)
    
    ' Brute space
    bruteSpace = wdDoc.PageSetup.pageHeight - markerPosition
    
    If first = True Then
        ' EffectiveSpace
        availableSpace = bruteSpace - wdDoc.PageSetup.BottomMargin '- wdDoc.PageSetup.TopMargin
        
    Else
        ' EffectiveSpace
        availableSpace = bruteSpace
    
    End If
        
    
    ' Pad height on the Excel
    rowHeightExcel = excelRange.Worksheet.Rows(1).Height
    
    
    CalculateRowsPerSlice = CInt(availableSpace / rowHeightExcel)
    
    

End Function

Sub MoveCursor(ByRef wdbmRange As Word.Range)
' Subrotina para mover o cursor de lugar

With wdbmRange
        .Collapse Direction:=wdCollapseEnd
        .InsertBreak Type:=wdPageBreak
        .Collapse Direction:=wdCollapseEnd ' Ajusta a posi��o do intervalo para o final da quebra.
    End With

End Sub



Function getWordDocument(choose As String, stWordReport As String, mark As String) As Object
    'Rotina para pegar o documento Word
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    

    'Initialize the Word objects.
    If choose = "Sim" Then
        Set wdApp = GetObject(, "Word.Application")
        Set wdDoc = wdApp.ActiveDocument
        
        With wdApp.Selection
            wdDoc.Bookmarks.Add Name:=mark, Range:=.Range
        End With
        
        
    Else
        ' Reference a words instance
        Set wdApp = New Word.Application
    
        Set wdDoc = wdApp.Documents.Open(ThisWorkbook.Path & "\" & stWordReport)
        
        
    End If
    
    Set getWordDocument = wdDoc
    
End Function

Sub signIgm(ByRef wdDoc As Object, versao As Integer, userName As String)
    ' Rotina para assinar imagens
    
    Dim metadados As String
    Dim imgRange As Word.Range
    Dim img As InlineShape

    metadata = "Criado por " & userName & " em " & Format(Date, "dd/mm/yyyy") & "Plan Version:" & versao

    ' Assinando a imagem
    Set img = wdDoc.InlineShapes(wdDoc.InlineShapes.Count)
            
    ' Usar o Range da imagem para adicionar o coment�rio
    img.AlternativeText = metadata


End Sub

Sub finalizeWordDocument(choose As String, ByRef wdDoc As Object, ByRef wdbmRange As Object)
' Rotina para finalizar o programa
        
        If choose = "N�o" Then
            With wdDoc
                .Save
                .Close
            End With
        
        Else
            With wdDoc
                .Save
            End With
        End If
        
        'Quit Word.
        If choose = "N�o" Then
            wdApp.Quit
        End If
        
        'Null out your variables.
        Set wdbmRange = Nothing
        Set wdDoc = Nothing
        Set wdApp = Nothing
        
        'Clear out the clipboard, and turn screen updating back on.
        With Application
            .CutCopyMode = False
            .ScreenUpdating = True
        End With
        
        MsgBox "Transfer�ncia realizado com sucesso " & vbNewLine & _
               "para " & stWordReport, vbInformation


End Sub
