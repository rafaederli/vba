Attribute VB_Name = "Módulo1"
Sub Fill()
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim pptSlide As Object
    Dim excelApp As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim slideIndex As Integer
    
    ' Define o número do slide para colar o gráfico (alterar conforme necessário)
    slideIndex = 1  ' Exemplo: primeiro slide
    
    ' Iniciar o Excel e abrir a planilha
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False ' Mantenha o Excel invisível
    Set wb = excelApp.Workbooks.Open("C:\Users\RAFAEL\Documents\repositories\vba\base.xlsx")
    
    ' Selecionar a aba e o gráfico desejado
    Set ws = wb.Sheets("compiled") ' Alterar para o nome da aba com o gráfico
    Set chartObj = ws.ChartObjects("Gráfico 3") ' Alterar para o nome do gráfico
    
    ' Copiar o gráfico
    chartObj.Copy
    
    ' Iniciar o PowerPoint e abrir a apresentação
    Set pptApp = CreateObject("PowerPoint.Application")
    Set pptPresentation = pptApp.Presentations.Open("C:\Users\RAFAEL\Documents\repositories\vba\presentation.pptx")
    
    ' Selecionar o slide em que o gráfico será colado
    Set pptSlide = pptPresentation.Slides(slideIndex)
    
    ' Colar o gráfico no slide selecionado e capturar a referência do ShapeRange
    Set ShapeRange = pptSlide.Shapes.PasteSpecial(DataType:=ppPasteEnhancedMetafile)
    
    ' Verifica se a colagem retornou uma coleção de objetos
    If Not ShapeRange Is Nothing Then
        ' Definir a posição e o tamanho do gráfico
        With ShapeRange(1) ' Trabalhar com o primeiro item da coleção
            .Left = 0  ' Distância do lado esquerdo (em pontos)
            .Top = 0 ' Distância do topo (em pontos)
        End With
    End If
    
    ' Salvar e fechar a apresentação
    pptPresentation.Save
    pptPresentation.Close
    
    ' Fechar o Excel
    wb.Close SaveChanges:=False
    excelApp.Quit
    pptApp.Quit
    
    ' Limpeza de objetos
    Set chartObj = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set excelApp = Nothing
    Set pptSlide = Nothing
    Set pptPresentation = Nothing
    Set pptApp = Nothing
    
    'MsgBox "Gráfico copiado com sucesso!"
End Sub

