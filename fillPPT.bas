Attribute VB_Name = "M�dulo1"
Sub Fill()
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim pptSlide As Object
    Dim excelApp As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim slideIndex As Integer
    
    ' Define o n�mero do slide para colar o gr�fico (alterar conforme necess�rio)
    slideIndex = 1  ' Exemplo: primeiro slide
    
    ' Iniciar o Excel e abrir a planilha
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False ' Mantenha o Excel invis�vel
    Set wb = excelApp.Workbooks.Open("C:\Users\RAFAEL\Documents\repositories\vba\base.xlsx")
    
    ' Selecionar a aba e o gr�fico desejado
    Set ws = wb.Sheets("compiled") ' Alterar para o nome da aba com o gr�fico
    Set chartObj = ws.ChartObjects("Gr�fico 3") ' Alterar para o nome do gr�fico
    
    ' Copiar o gr�fico
    chartObj.Copy
    
    ' Iniciar o PowerPoint e abrir a apresenta��o
    Set pptApp = CreateObject("PowerPoint.Application")
    Set pptPresentation = pptApp.Presentations.Open("C:\Users\RAFAEL\Documents\repositories\vba\presentation.pptx")
    
    ' Selecionar o slide em que o gr�fico ser� colado
    Set pptSlide = pptPresentation.Slides(slideIndex)
    
    ' Colar o gr�fico no slide selecionado e capturar a refer�ncia do ShapeRange
    Set ShapeRange = pptSlide.Shapes.PasteSpecial(DataType:=ppPasteEnhancedMetafile)
    
    ' Verifica se a colagem retornou uma cole��o de objetos
    If Not ShapeRange Is Nothing Then
        ' Definir a posi��o e o tamanho do gr�fico
        With ShapeRange(1) ' Trabalhar com o primeiro item da cole��o
            .Left = 0  ' Dist�ncia do lado esquerdo (em pontos)
            .Top = 0 ' Dist�ncia do topo (em pontos)
        End With
    End If
    
    ' Salvar e fechar a apresenta��o
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
    
    'MsgBox "Gr�fico copiado com sucesso!"
End Sub

