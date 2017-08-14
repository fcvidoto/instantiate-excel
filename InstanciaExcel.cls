VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstanciaExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private m_xls As Object 'Excel.Application
Private m_workbookName As String
Private m_sheetName As String
Private m_workbookInstanciado As Object
Private m_meWBK As Object

'Instancia do Excel 'Excel.Application'
Public Property Get xls() As Object: Set xls = m_xls: End Property
Public Property Set xls(ByVal valor As Object): Set m_xls = valor: End Property

'Nome do workbook que será utilizado
Public Property Get workbookName() As String: workbookName = m_workbookName: End Property
Public Property Let workbookName(ByVal valor As String): m_workbookName = valor: End Property

'Sheet que será usada na instancia
Public Property Get sheetName() As String: sheetName = m_sheetName: End Property
Public Property Let sheetName(ByVal valor As String): m_sheetName = valor: End Property

'Caminho que o workbook foi salvo
Public Property Get workbookFullpath() As String: workbookFullpath = Me.workbookInstanciado.FullName: End Property

'Workbook instanciado
Public Property Get workbookInstanciado() As Object: Set workbookInstanciado = m_workbookInstanciado: End Property
Public Property Set workbookInstanciado(ByVal valor As Object): Set m_workbookInstanciado = valor: End Property

'O primeiro workbook
Public Property Get meWBK() As Object: Set meWBK = m_meWBK: End Property
Public Property Set meWBK(ByVal valor As Object): Set m_meWBK = valor: End Property

'Cria um novo Excel
Sub NewFile()
  Set Me.xls = createObject("Excel.Application") 'New Excel.Application
  Me.xls.Workbooks.Add
End Sub

'Salva arquivo no endereco
Sub SaveAs(endereco As String)
  Me.xls.ActiveWorkbook.SaveAs endereco
End Sub

'Se a sheet não existe na RQ, cria a sheet
Function SeSheetNaoExisteCriaSheet(ByRef sheetAondeCola As String) As Boolean
        Dim sheet As Object
        Dim validacao As Boolean
        validacao = False
        '----------------------------'
        'Itera por todas as sheets do Excel e verifica se a sheet não existe
        For Each sheet In Me.workbookInstanciado.Sheets
        '----------------------------'
                'Verifica se a sheet existe
                If sheet.name = sheetAondeCola Then
            SeSheetNaoExisteCriaSheet = True
            Exit Function
                End If
        Next sheet
End Function

'cria a tabela no range de dados
Public Sub converteRangeTabela(nomeDaSheet As String, _
                              nomeDoReport As String, _
                              is_tabelaComum As Boolean, _
                              colunaLetraInicial As Variant, _
                              colunaNumeroInicial As Variant, _
                              isExcecao As Boolean)
  If is_tabelaComum = False Then
    xls.Sheets(nomeDaSheet).ListObjects.Add(1, xls.Sheets(nomeDaSheet).UsedRange, , 1).name = nomeDoReport
    xls.Sheets(nomeDaSheet).UsedRange.Select
    xls.Sheets(nomeDaSheet).ListObjects(nomeDoReport).TableStyle = ""
    Exit Sub
  End If
  '-----------------------------------------------------------------
  'comeca com a coluna A
  Dim ultimaLinha As Double
  ultimaLinha = xls.Sheets(nomeDaSheet).Range(colunaLetraInicial & xls.Sheets(nomeDaSheet).Rows.Count).End(-4162).Row
  Dim primeiraLinha As Double
  '**********************************
  'se for da coluna 'A' nao soma o +1
  If colunaLetraInicial = "A" Then
    primeiraLinha = xls.Sheets(nomeDaSheet).Range(colunaLetraInicial & ultimaLinha).End(-4162).Row  'nao pega o titulo do report
  Else
    primeiraLinha = xls.Sheets(nomeDaSheet).Range(colunaLetraInicial & ultimaLinha).End(-4162).Row + 1  'nao pega o titulo do report
  End If
  '**********************************
  'descobre a ultima coluna
  Dim ultimaColuna As Double
  ultimaColuna = xls.Sheets(nomeDaSheet).Range(colunaLetraInicial & primeiraLinha).End(-4161).Column
  'seleciona o range da tabela
  Dim enderecoRange As String
  enderecoRange = xls.Sheets(nomeDaSheet).Range(xls.Sheets(nomeDaSheet).Cells(primeiraLinha, CDbl(colunaNumeroInicial)), xls.Sheets(nomeDaSheet).Cells(ultimaLinha + IIf(isExcecao, 1, 0), ultimaColuna)).Address
On Error Resume Next
  xls.Sheets(nomeDaSheet).ListObjects.Add(1, xls.Sheets(nomeDaSheet).Range(enderecoRange), , 1).name = nomeDoReport
  xls.Sheets(nomeDaSheet).Range(enderecoRange).Select
  xls.Sheets(nomeDaSheet).ListObjects(nomeDoReport).TableStyle = ""
End Sub

'cria a tabela no range de dados
Public Sub converteRangeTabela_Scorecards(nomeDaSheet As String, _
                                          nomeDoReport As String, _
                                          is_tabelaComum As Boolean, _
                                          colunaLetraInicial As Variant, _
                                          colunaNumeroInicial As Variant, _
                                          isExcecao As Boolean)
  If is_tabelaComum = False Then
    xls.Sheets(nomeDaSheet).ListObjects.Add(1, xls.Sheets(nomeDaSheet).UsedRange, , 1).name = nomeDoReport
    xls.Sheets(nomeDaSheet).UsedRange.Select
    xls.Sheets(nomeDaSheet).ListObjects(nomeDoReport).TableStyle = ""
    Exit Sub
  End If
  '-----------------------------------------------------------------
  'comeca com a coluna A
  Dim ultimaLinha As Double
  ultimaLinha = xls.Sheets(nomeDaSheet).Range(colunaLetraInicial & xls.Sheets(nomeDaSheet).Rows.Count).End(-4162).Row
  Dim primeiraLinha As Double
  primeiraLinha = xls.Sheets(nomeDaSheet).Range(colunaLetraInicial & ultimaLinha).End(-4162).Row 'nao pega o titulo do report
  Dim ultimaColuna As Double 'descobre a ultima coluna
  ultimaColuna = xls.Sheets(nomeDaSheet).Range(colunaLetraInicial & primeiraLinha).End(-4161).Column
  '-----------------------------------------------------------------
  'seleciona o range da tabela
  Dim enderecoRange As String
  enderecoRange = xls.Sheets(nomeDaSheet).Range(xls.Sheets(nomeDaSheet).Cells(primeiraLinha, CDbl(colunaNumeroInicial)), xls.Sheets(nomeDaSheet).Cells(ultimaLinha + IIf(isExcecao, 1, 0), ultimaColuna)).Address
  '-----------------------------------------------------------------
On Error Resume Next
  xls.Sheets(nomeDaSheet).ListObjects.Add(1, xls.Sheets(nomeDaSheet).Range(enderecoRange), , 1).name = nomeDoReport
  xls.Sheets(nomeDaSheet).Range(enderecoRange).Select
  xls.Sheets(nomeDaSheet).ListObjects(nomeDoReport).TableStyle = ""
End Sub

'Muda a config do Excel para o tamanho 80
Public Sub ZoomOitenta(nomeDaSheet As String)
  Me.xls.Sheets(nomeDaSheet).Select 'seleciona a sheet
  Me.xls.ActiveWindow.Zoom = 80
End Sub

'formata as casas decimais da sheet
Public Sub formataCasasDecimaisDaSheet(sheetName As String)
  Me.xls.Sheets(sheetName).UsedRange.NumberFormat = "0.000"
End Sub

'formata a primeira coluna da sheet como data
'Public Sub formataPrimeiraColunaData(nomeDaSheet As String)
'  Me.xls.Sheets(nomeDaSheet).Columns("A:A").NumberFormat = "dd/mm/yyyy"
'End Sub

'insere titulo do relatorio
Public Sub insereTitulo(sheetName As String, titulo As String, tituloEspecial As Boolean)
  Dim ultimaLinha As Double
  ultimaLinha = xls.Sheets(sheetName).Range("B" & xls.Sheets(sheetName).Rows.Count).End(-4162).Row + IIf(tituloEspecial, 1, 0) + 1
  xls.Sheets(sheetName).Range("B" & ultimaLinha + 2).Value = titulo
  '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  'altera o fundo e negrita
  xls.Sheets(sheetName).Range(xls.Sheets(sheetName).Cells(ultimaLinha + 2, 2), xls.Sheets(sheetName).Cells(ultimaLinha + 2, 3)).Font.Bold = True
  If tituloEspecial Then
    xls.Sheets(sheetName).Range(xls.Sheets(sheetName).Cells(ultimaLinha + 2, 2), xls.Sheets(sheetName).Cells(ultimaLinha + 2, 6 - IIf(tituloEspecial, 2, 0))).Interior.Color = 0 'rgbBlack
    xls.Sheets(sheetName).Range(xls.Sheets(sheetName).Cells(ultimaLinha + 2, 2), xls.Sheets(sheetName).Cells(ultimaLinha + 2, 3)).Font.Color = 16777215  'rgbWhite
  Else
    xls.Sheets(sheetName).Range(xls.Sheets(sheetName).Cells(ultimaLinha + 2, 2), xls.Sheets(sheetName).Cells(ultimaLinha + 2, 6)).Interior.Color = 13882323 'rgbLightGrey
  End If
End Sub

'insere titulo do relatorio
Public Sub insereTituloScorecards(sheetName As String, titulo As String, tituloEspecial As Boolean, colunaLetra As String, colunaNumero As Variant)
  Dim ultimaLinha As Double
  ultimaLinha = xls.Sheets(sheetName).Range(colunaLetra & xls.Sheets(sheetName).Rows.Count).End(-4162).Row + 2  'coluna 'B' fixa
  xls.Sheets(sheetName).Range(colunaLetra & ultimaLinha).Value = titulo
  '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  'altera o fundo e negrita
  If tituloEspecial Then
    xls.Sheets(sheetName).Range(colunaLetra & ultimaLinha).Font.Bold = True
    xls.Sheets(sheetName).Range(xls.Sheets(sheetName).Cells(ultimaLinha, CDbl(colunaNumero)), xls.Sheets(sheetName).Cells(ultimaLinha, CDbl(colunaNumero) + 1)).Interior.Color = 0 'rgbBlack
    xls.Sheets(sheetName).Range(xls.Sheets(sheetName).Cells(ultimaLinha, CDbl(colunaNumero)), xls.Sheets(sheetName).Cells(ultimaLinha, CDbl(colunaNumero) + 1)).Font.Color = 16777215  'rgbWhite
  Else
    xls.Sheets(sheetName).Range(colunaLetra & ultimaLinha).Font.Bold = True
    xls.Sheets(sheetName).Range(xls.Sheets(sheetName).Cells(ultimaLinha, CDbl(colunaNumero)), xls.Sheets(sheetName).Cells(ultimaLinha, CDbl(colunaNumero) + 1)).Interior.Color = 13882323 'rgbLightGrey
  End If
End Sub

'cola os dados do recordset na planilha | formata o estilo da header aqui
Sub ColaDadosRS(ByVal rs As Recordset, _
                      sheetName As String, _
                      colunaLetraInicial As Variant, _
                      colunaNumeroInicial As Variant)
  Dim i As Integer
  Dim ultimaLinha As Double
  Dim x As Integer
  Dim k As Integer
  x = colunaNumeroInicial
  k = colunaNumeroInicial
  '--------------------------------------------------------------
  ultimaLinha = xls.Sheets(sheetName).Range("B" & xls.Sheets(sheetName).Rows.Count).End(-4162).Row  'coluna 'B' fixa
  'se for o primeiro report, comeca na linha 2
  If ultimaLinha = 1 Then
    ultimaLinha = 3
  Else
    ultimaLinha = ultimaLinha + 1
  End If
  '--------------------------------------------------------------
  'novas alteraçoes, coluna inicial variável <-----------------
  For i = 1 To rs.Fields.Count
    Me.xls.Sheets(sheetName).Cells(ultimaLinha, k) = rs.Fields(i - 1).name
    k = k + 1
  Next i
  '--------------------------------------------------------------
  rs.MoveFirst
  Me.xls.Sheets(sheetName).Range(colunaLetraInicial & ultimaLinha + 1).CopyFromRecordset rs
End Sub

Sub ColaDadosRS_Scorecards(ByVal rs As Recordset, _
                          sheetName As String, _
                          colunaLetraInicial As Variant, _
                          colunaNumeroInicial As Variant)
  Dim i As Integer
  Dim ultimaLinha As Double
  Dim x As Integer
  Dim k As Integer
  x = colunaNumeroInicial
  k = colunaNumeroInicial
  '--------------------------------------------------------------
  ultimaLinha = xls.Sheets(sheetName).Range(colunaLetraInicial & xls.Sheets(sheetName).Rows.Count).End(-4162).Row + 1  'coluna 'B' fixa
  'se for o primeiro report, comeca na linha 2
  If ultimaLinha = 1 Then
    ultimaLinha = 3
  Else
    ultimaLinha = ultimaLinha + 1
  End If
  '--------------------------------------------------------------
  'novas alteraçoes, coluna inicial variável <-----------------
  For i = 1 To rs.Fields.Count
    Me.xls.Sheets(sheetName).Cells(ultimaLinha, k) = rs.Fields(i - 1).name
    k = k + 1
  Next i
  rs.MoveFirst
  Me.xls.Sheets(sheetName).Range(colunaLetraInicial & ultimaLinha + 1).CopyFromRecordset rs
End Sub

'Captura a isntancia do Excel
Private Sub GetInstance()
  On Error GoTo createNewInstance
  Set xls = GetObject(, "Excel.Application")
Exit Sub
'Se der erro na captura da instancia, cria uma nova
createNewInstance:
  Set xls = createObject("Excel.Application") 'New Excel.Application
End Sub

'Captura dados da instância do Excel
Public Function InstanciaObjeto(ByVal endereco As String, _
                       Optional ByRef meWBK As Object, _
                       Optional ByRef sheet As String) As Boolean
  'On Error GoTo ErroEncontrado
  'Captura a isntancia do Excel
  Call GetInstance
  Dim xlsName As String
  xlsName = Me.NomeDoArquivo(endereco) 'Pega apenas o nome do arquivo no endereço utilizado
  Me.FechaWorkbook (xlsName) 'Fecha e salva a instancia caso ela esteja aberta
  '------------------------------------------------------------------
  'Abre a instancia do Excel para pegar o arquivo de clientes
  With xls
    .DisplayAlerts = False
    .EnableEvents = False
    .Workbooks.Open endereco, UpdateLinks:=False, ReadOnly:=False
    Set Me.workbookInstanciado = Me.SetaWorkbook(endereco, Me) 'Seta o workbook na propriedade da classe
  End With
  '------------------------------------------------------------------
  If Not meWBK Is Nothing Then 'Volta pra janela inicial
    Set Me.meWBK = meWBK
    Me.meWBK.Activate
  End If
  '------------------------------------------------------------------
  If sheet = vbNullString Then 'Define o nome da Sheet (Se for vazio, pega a primeira sheetName
    sheet = workbookInstanciado.Sheets(1).name
    sheetName = sheet
  End If
  InstanciaObjeto = True 'Instanciou o obj na memoria
Exit Function
ErroEncontrado:
        InstanciaObjeto = False
End Function

'Mostra todos os dados da sheet
Private Sub ShowAllData(ByVal sheetDaBase As String)
On Error Resume Next
    Me.workbookInstanciado.Sheets(sheetDaBase).Cells.AutoFilter
End Sub

'Captura o array de dados do excel instanciado
Public Function CapturaDados(ByVal sheetDaBase As String) As Variant
  Call ShowAllData(sheetDaBase) 'Remove qualquer filtro da sheet
  Dim ultimaColuna As String
        Dim DataRange As Variant
        Dim rangeAddress As String
  Dim linha As Long
  Dim ultimaLinha As String
  Dim j As Long
        '-----------------------------------
  ultimaColuna = Me.workbookInstanciado.Sheets(sheetDaBase).Range("A1").End(-4161).Address '(Capturando range da tabela e montanto o array na memória)
  ultimaLinha = Me.workbookInstanciado.Sheets(sheetDaBase).Cells(linha, Me.workbookInstanciado.Sheets(sheetDaBase).Range("A1").End(-4161).Column).Address
  linha = Me.workbookInstanciado.Sheets(sheetDaBase).Range("A1").End(-4121).Row
        '-----------------------------------
  'Validando tamanho do array (Se for fora dos parametros, retorna erro de validação)
  If Me.workbookInstanciado.Sheets(sheetDaBase).ProtectContents _
                          Or ultimaColuna = "$XFD$1" _
                          Or ultimaColuna = "$IV$1" _
                          Or ultimaLinha = "$XFD$1048576" Then
              
              Exit Function
  End If
        '-----------------------------------
        'O endereço do range que será capturado
        rangeAddress = "A1:" & ultimaLinha
        Me.workbookInstanciado.Sheets(sheetDaBase).Range(rangeAddress).NumberFormat = "General" 'Corrigindo bug
        DataRange = Me.workbookInstanciado.Sheets(sheetDaBase).Range(rangeAddress).Value 'Array de dados
        CapturaDados = DataRange
End Function

'Fecha o Excel quando acabar a instância
Private Sub Class_Terminate()
        If Not Me.xls Is Nothing Then
          Me.xls.ActiveWorkbook.Close False
        End If
End Sub

'Valida se a sheet existe no excel Instanciado
Public Function ValidaSeSheetExiste(ByVal nomeDaSheet As String) As Boolean
  Dim sheet As Object
  ValidaSeSheetExiste = False 'Valor default da function
  For Each sheet In Me.xls.Sheets 'Itera pela lista de sheets
    If sheet.name = nomeDaSheet Then 'Se a sheet existe, retorna True
      ValidaSeSheetExiste = True
      Exit Function
    End If   '------------------------'
  Next sheet
End Function

'Valida se a quantidade de colunas está de acordo com a validação
Public Function ValidaQtdColunas(ByVal rangeDeDados As Variant, _
                                 ByRef qtdDeColunasNecessaria) As Boolean
  If UBound(rangeDeDados, 2) = qtdDeColunasNecessaria Then 'Se o número de colunas bater com q qtd certa
    ValidaQtdColunas = True
  Else
    ValidaQtdColunas = False
        End If
End Function

'Salva a instancia na maquina
Public Sub Salva(ByVal pathBase As String)
  workbookInstanciado.Save
End Sub

'Valida se a header da instancia de RQ possui os mesmos parametros
Public Function HeaderError(ByVal sheetAondeCola As String, _
                            ByVal headerArray As Variant) As Boolean

    Dim ultimaColuna As String: ultimaColuna = Me.workbookInstanciado.Sheets(sheetAondeCola).Range("A1").End(-4161).Address
    Dim linha As Long: linha = Me.workbookInstanciado.Sheets(sheetAondeCola).Range("A1").End(-4121).Row
    Dim ultimaLinha As String: ultimaLinha = Me.workbookInstanciado.Sheets(sheetAondeCola).Cells(linha, Me.workbookInstanciado.Sheets(sheetAondeCola).Range("A1").End(-4161).Column).Address
    Dim j As Long
    
    'Validando tamanho do array (Se for fora dos parametros, retorna erro de validação)
    If Me.workbookInstanciado.Sheets(sheetAondeCola).ProtectContents _
                            Or ultimaColuna = "$XFD$1" _
                            Or ultimaColuna = "$IV$1" _
                            Or ultimaLinha = "$XFD$1048576" Then
                
                HeaderError = True
                Exit Function
        
    End If
    
    'O endereço do range que será utilizado
    Dim rangeAddress As String: rangeAddress = "A1:" & ultimaColuna
    
    Dim i As Integer
    
    'Instancia só a header de sheetAondeCola
    Dim dataRangeHeader As Variant: dataRangeHeader = Me.workbookInstanciado.Sheets(sheetAondeCola).Range(rangeAddress).Value
    
    'Itera por cada valor da header instanciada comparando com os parametros de headerArray
    For i = 1 To UBound(headerArray, 2)
              
        'Verifica um a um se os dados da header estão batendo
        If UCase(dataRangeHeader(1, i)) <> UCase(headerArray(1, i)) Then

            'Erro nos dados da header
            HeaderError = True
            Exit Function
        End If
    Next i

End Function

'Copia todos os nomes do workbook
Public Sub CopiaNames(nomes() As Object)
  Dim name As Object
  Dim i As Integer: i = UBound(nomes)
  For Each name In Me.workbookInstanciado.Names 'Itera por todos os nomes do workbook e cria um array de nomes
    i = i + 1 'Redimensiona o array de acordo com o array de nomes
    ReDim Preserve nomes(1 To i)
    Set nomes(i) = name
  Next name
End Sub

'Captura os dados da Header
Public Function MontaHeaderArray(ByVal sheetDaBaseDeDados As String) As Variant
        Dim ultimaColuna As String
        Dim linha As Long
        Dim ultimaLinha As String
        Dim rangeAddress As String
        Dim dataRangeHeader As Variant
        Dim j As Long
        Dim i As Integer
        '-------------------------------'
        ultimaColuna = Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range("A1").End(-4161).Address
        linha = Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range("A1").End(-4121).Row
        ultimaLinha = Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Cells(linha, Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range("A1").End(-4161).Column).Address
        '-------------------------------'
        'Validando tamanho do array (Se for fora dos parametros, retorna erro de validação)
        If Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).ProtectContents _
                                Or ultimaColuna = "$XFD$1" _
                                Or ultimaColuna = "$IV$1" _
                                Or ultimaLinha = "$XFD$1048576" Then
                    
                    Exit Function
        End If
        '-------------------------------'
        rangeAddress = "A1:" & ultimaColuna 'O endereço do range que será utilizado
        dataRangeHeader = Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range(rangeAddress).Value 'Instancia só a header de sheetDaBaseDeDados
        MontaHeaderArray = dataRangeHeader 'Retorna o array da header
End Function

'Retorna a lista de sheets do arquivo
Public Function SheetsDoArquivo() As Variant
  Dim sheet As Object
  Dim lista() As Variant
  Dim i As Integer
  '--------------------------------'
  'Itera pela lista de sheets
  For Each sheet In Me.workbookInstanciado.Worksheets
    i = i + 1
    ReDim Preserve lista(1 To i)
    '--------------------------------'
    If Me.workbookInstanciado.ActiveSheet.name = sheet.name Then 'Adiciona cada sheet na listbox
      lista(i) = "*" & sheet.name
    Else
      lista(i) = sheet.name
    End If
  Next sheet
  SheetsDoArquivo = lista 'Retorna a lista de sheets do arquivo
End Function

'Desprotege todas as planilhas do workbook
Public Sub DesprotegePlanilhas()
  Dim sheet As Object
  For Each sheet In Me.workbookInstanciado.Sheets 'Itera por todas as sheets do workbook desprotegendo
  Next sheet
End Sub

'Protege todas as planilhas do workbook
Public Sub ProtegePlanilhas()
  Dim sheet As Object
  For Each sheet In Me.workbookInstanciado.Sheets
    'Itera por todas as sheets do workbook desprotegendo
  Next sheet
End Sub

'Cria a sheet no instancia
Sub CriaSheet(ByRef sheetAondeCola As String)
On Error GoTo exitSub
  Dim qtdSheets As Integer
  qtdSheets = Me.workbookInstanciado.Sheets.Count
  Me.workbookInstanciado.Sheets.Add After:=Me.workbookInstanciado.Sheets(qtdSheets) 'Adiciona a sheet como primeira sheet
  Me.workbookInstanciado.Sheets(qtdSheets + 1).name = sheetAondeCola 'Renomeia a sheet adicionada
exitSub:
End Sub

'cria a quantidade de sheets necessarias no arquivo
Public Sub CriaQTDSheet(qtdDeSheets As Integer)
  While Not Me.xls.Sheets.Count = qtdDeSheets
    Me.xls.Sheets.Add
  Wend
End Sub

'verifica se exitem dados na sheet
Public Function existeDadosSheet(sheetName As String, _
                                 linhaInicial As Integer, _
                                 colunaInicial As Integer) As Boolean
  Dim ultimaLinha As Double
  ultimaLinha = xls.Sheets(sheetName).Cells(xls.Sheets(sheetName).Rows.Count, 2).End(-4162).Row
  If ultimaLinha < linhaInicial Then 'verifique se existe dados na sheet
    existeDadosSheet = False
  End If
  existeDadosSheet = True
End Function


'Desprotege a planilha
'Sub DesprotegerPlanilha(ByRef sheet As Object)
'  Dim i, j, k, l, m, i1, i2, i3, i4, i5, i6, n As Integer
'On Error Resume Next
'  For i = 65 To 66
'    For j = 65 To 66
'      For k = 65 To 66
'        For l = 65 To 66
'          For m = 65 To 66
'            For i1 = 65 To 66
'              For i2 = 65 To 66
'                For i3 = 65 To 66
'                  For i4 = 65 To 66
'                    For i5 = 65 To 66
'                      For i6 = 65 To 66
'                        For n = 32 To 126
'                            sheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
'
'                            If ActiveSheet.ProtectContents = False Then
'                                'MsgBox "A planilha foi desprotegida com sucesso!!!"
'                                Exit Sub
'                            End If
'
'                        Next
'                      Next
'                    Next
'                  Next
'                Next
'              Next
'            Next
'          Next
'        Next
'      Next
'    Next
'  Next
'End Sub

'fecha a instancia caso esteja aberta'
Public Sub FechaWorkbook(ByVal nomeDoWorkbookComExtensao As String)
  Dim ArquivosAbertos As Object
  Dim xls As Object
  Set xls = GetObject(, "Excel.Application")
  For Each ArquivosAbertos In xls.Workbooks 'Itera por todos os workbooks abertos
  'Fecha o workbook específico E não salva as alterações
  If ArquivosAbertos.name = nomeDoWorkbookComExtensao Or _
    ArquivosAbertos.name = "formulario.xlsx" Then
    ArquivosAbertos.Close False
    Exit Sub
  End If
  '---------------------'
  Next ArquivosAbertos
End Sub


'Retorna o nome do arquivo a partir do endereço
Public Function NomeDoArquivo(ByVal pathEndereco As String) As String
  Dim fso As Object
  Dim pasta As Object
  Dim arquivo As Object
  '------------------------'
  Set fso = createObject("Scripting.FileSystemObject") 'New Scripting.FileSystemObject
  Set arquivo = fso.GetFile(pathEndereco) 'Set o arquivo do endereço
  NomeDoArquivo = arquivo.name 'Retorna o nome do arquivo
End Function

'Seta o workbook na propriedade da classe
Public Function SetaWorkbook(ByVal pathEndereco As String, _
                             ByRef InstanciaExcel As InstanciaExcel) As Object
  Dim fso As Object
  Set fso = createObject("Scripting.FileSystemObject")
  Dim pasta As Object
  Dim arquivo As Object
  Dim Workbook As Object
  Dim fileName As String
  Set arquivo = fso.GetFile(pathEndereco) 'Set o arquivo do endereço
  fileName = arquivo.name 'Nome do workbook
  '-----------------'
  For Each Workbook In InstanciaExcel.xls.Workbooks 'Procura pelo workbook na instancia
    '-----------------'
    If Workbook.name = fileName Then 'Se achar o workbook em algum workbook
      Set SetaWorkbook = Workbook 'Return
      Exit Function
    End If
  Next Workbook
End Function


