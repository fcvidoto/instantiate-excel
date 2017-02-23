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

Private m_xls As Excel.Application
Private m_workbookName As String
Private m_sheetName As String
Private m_workbookInstanciado As Workbook
Private m_meWBK As Object

'Instancia do Excel
Public Property Get xls() As Excel.Application: Set xls = m_xls: End Property
Public Property Set xls(ByVal valor As Excel.Application): Set m_xls = valor: End Property

'Nome do workbook que será utilizado
Public Property Get workbookName() As String: workbookName = m_workbookName: End Property
Public Property Let workbookName(ByVal valor As String): m_workbookName = valor: End Property

'Sheet que será usada na instancia
Public Property Get sheetName() As String: sheetName = m_sheetName: End Property
Public Property Let sheetName(ByVal valor As String): m_sheetName = valor: End Property

'Caminho que o workbook foi salvo
Public Property Get workbookFullpath() As String: workbookFullpath = Me.workbookInstanciado.FullName: End Property

'Workbook instanciado
Public Property Get workbookInstanciado() As Workbook: Set workbookInstanciado = m_workbookInstanciado: End Property
Public Property Set workbookInstanciado(ByVal valor As Workbook): Set m_workbookInstanciado = valor: End Property

'O primeiro workbook
Public Property Get meWBK() As Workbook: Set meWBK = m_meWBK: End Property
Public Property Set meWBK(ByVal valor As Object): Set m_meWBK = valor: End Property

'Cria um novo Excel
Sub NewFile()
  Set Me.xls = New Excel.Application
  Me.xls.Workbooks.Add
End Sub

'Salva arquivo no endereco
Sub SaveAs(endereco As String)
  Me.xls.ActiveWorkbook.SaveAs endereco
End Sub

'Se a sheet não existe na RQ, cria a sheet
Function SeSheetNaoExisteCriaSheet(ByRef sheetAondeCola As String) As Boolean
    
    Dim sheet As Worksheet
    Dim validacao As Boolean
        validacao = False
    
    'Itera por todas as sheets do Excel e verifica se a sheet não existe
    For Each sheet In Me.workbookInstanciado.Sheets
        
        'Verifica se a sheet existe
        If sheet.name = sheetAondeCola Then
            SeSheetNaoExisteCriaSheet = True
            Exit Function
        End If
    Next sheet
    
End Function

'cola os dados do recordset na planilha
Sub ColaDadosRS(ByVal rs As Recordset, Optional strSQL As String)
  
  For i = 0 To rs.Fields.Count - 1
    Me.xls.Sheets(1).Cells(1, i + 1) = rs.Fields(i).name
  Next i
  
  Me.xls.Sheets(1).Range("A2").CopyFromRecordset rs
End Sub

'Captura a isntancia do Excel
Private Sub GetInstance()

    On Error GoTo createNewInstance

    Set xls = GetObject(, "Excel.Application")

Exit Sub
'Se der erro na captura da instancia, cria uma nova
createNewInstance:
    
    Set xls = New Excel.Application

End Sub

'Captura dados da instância do Excel
Public Function InstanciaObjeto(ByVal endereco As String, _
                       Optional ByRef meWBK As Workbook, _
                       Optional ByRef sheet As String) As Boolean
       
    'On Error GoTo ErroEncontrado
    
    'Captura a isntancia do Excel
    Call GetInstance
    
    Dim xlsName As String
    
    'Fecha o workbook antes de instancia-lo
    Dim CriandoObjetos As CriandoObjetos: Set CriandoObjetos = New CriandoObjetos
    
    'Pega apenas o nome do arquivo no endereço utilizado
    xlsName = CriandoObjetos.NomeDoArquivo(endereco)
    
    'Fecha e salva a instancia caso ela esteja aberta
    Call CriandoObjetos.FechaWorkbook(xlsName)

'------------------------------------------------------------------
    'Abre a instancia do Excel para pegar o arquivo de clientes
    With xls
        .DisplayAlerts = False
        .EnableEvents = False
        .Workbooks.Open endereco, UpdateLinks:=False, ReadOnly:=False
        
        'Seta o workbook na propriedade da classe
        Set Me.workbookInstanciado = CriandoObjetos.SetaWorkbook(endereco, Me)
        
    End With

    'Volta pra janela inicial
    If Not meWBK Is Nothing Then
        Set Me.meWBK = meWBK
        Me.meWBK.Activate
    End If
    
    'Define o nome da Sheet (Se for vazio, pega a primeira sheetName
    If sheet = vbNullString Then
        sheet = workbookInstanciado.Sheets(1).name
        sheetName = sheet
    End If
    
    'Instanciou o obj na memoria
    InstanciaObjeto = True

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

    'Remove qualquer filtro da sheet
    Call ShowAllData(sheetDaBase)

    '(Capturando range da tabela e montanto o array na memória)
    Dim ultimaColuna As String: ultimaColuna = Me.workbookInstanciado.Sheets(sheetDaBase).Range("A1").End(xlToRight).Address
    Dim linha As Long: linha = Me.workbookInstanciado.Sheets(sheetDaBase).Range("A1").End(xlDown).Row
    Dim ultimaLinha As String: ultimaLinha = Me.workbookInstanciado.Sheets(sheetDaBase).Cells(linha, Me.workbookInstanciado.Sheets(sheetDaBase).Range("A1").End(xlToRight).Column).Address
    Dim j As Long
    
    'Validando tamanho do array (Se for fora dos parametros, retorna erro de validação)
    If Me.workbookInstanciado.Sheets(sheetDaBase).ProtectContents _
                            Or ultimaColuna = "$XFD$1" _
                            Or ultimaColuna = "$IV$1" _
                            Or ultimaLinha = "$XFD$1048576" Then
                
                Exit Function
        
    End If
    
    'O endereço do range que será capturado
    Dim rangeAddress As String: rangeAddress = "A1:" & ultimaLinha
    Dim dataRange As Variant

    Me.workbookInstanciado.Sheets(sheetDaBase).Range(rangeAddress).NumberFormat = "General" 'Corrigindo bug
    
    'Array de dados
    dataRange = Me.workbookInstanciado.Sheets(sheetDaBase).Range(rangeAddress).Value
    
    CapturaDados = dataRange

End Function

'Fecha o Excel quando acabar a instância
Private Sub Class_Terminate()
    
    If Not Me.workbookInstanciado Is Nothing Then
        Me.workbookInstanciado.Close False
    End If

End Sub

'Valida se a sheet existe no excel Instanciado
Public Function ValidaSeSheetExiste(ByVal nomeDaSheet As String) As Boolean
    
    Dim sheet As Worksheet
    
    'Valor default da function
    ValidaSeSheetExiste = False
    
    'Itera pela lista de sheets
    For Each sheet In Me.workbookInstanciado.Worksheets
        
        'Se a sheet existe, retorna True
        If sheet.name = nomeDaSheet Then
            
            ValidaSeSheetExiste = True
            Exit Function
            
        End If
    Next sheet
    
End Function

'Valida se o nome das colunas está de acordo com a validação
Public Function ValidaNomeDasColunas(ByVal rangeDeDados As Variant, _
                                     ByVal arrayDeColunas As Variant, _
                                     Optional ByRef planilha As String) As Boolean

    Dim i As Integer

    'Valor default da validação de nomes
    ValidaNomeDasColunas = True
    
    'Verifica se é qualquer outra planilha, ou se pertence é planilha Preço NET
    'Aqui entra todas as planilhas, menos a PrecoNet
    If planilha <> "PrecoNET" Then
    
        'Itera pelo cabeçalho do datarange
        For i = 1 To UBound(rangeDeDados, 2)
            
            'Se o nome da coluna não for igual, retorna 'false' (erro)
            If rangeDeDados(1, i) <> arrayDeColunas(i) Then
                ValidaNomeDasColunas = False
                Exit Function
            End If
        Next i
    
    'Se for Preço net, tem um esquema próprio de validação
    Else
        'Itera pelo cabeçalho do datarange
        For i = 1 To UBound(rangeDeDados, 2)
            
            If i <= 2 Then
            'Se o nome da coluna não for igual, retorna 'false' (erro)
            If rangeDeDados(1, i) <> arrayDeColunas(i) Then
                ValidaNomeDasColunas = False
                Exit Function
            End If
            
            'Começa as validações de data
            Else
                If CDate(rangeDeDados(1, i)) <> arrayDeColunas(i) Then
                    ValidaNomeDasColunas = False
                    Exit Function
                End If
            End If
        Next i
    End If

End Function

'Valida se a quantidade de colunas está de acordo com a validação
Public Function ValidaQtdColunas(ByVal rangeDeDados As Variant, _
                                 ByRef qtdDeColunasNecessaria) As Boolean

    'Se o número de colunas bater com q qtd certa
    If UBound(rangeDeDados, 2) = qtdDeColunasNecessaria Then
        ValidaQtdColunas = True
    Else
        ValidaQtdColunas = False
    End If

End Function

'Cola os dados do array na planilha
Public Sub ColaOsDados(ByVal sheetName As String, _
                       ByVal dados As Variant, _
                       ByVal pathBase As String, _
                       ByVal apagarDadosAnteriores As Boolean)
    
    Dim linhaInicial As Integer: linhaInicial = 1 'Começa a colar os dados na linha 2
    Dim linhaFinal As Integer: linhaFinal = UBound(dados)
    Dim colunaInicial As Integer: colunaInicial = 1
    Dim colunaFinal As Integer: colunaFinal = UBound(dados, 2)
    Dim rangeAddress As String
    
    'Não cola na linha 1, começa na linha 2
    If linhaInicial <> 1 Then
        linhaFinal = linhaFinal + 1
    End If
    
    'Se não houver workbook, cria um novo só para calcular o address (bug do Excel)
    If Application.Workbooks.Count = 0 Then
        Workbooks.Add
    End If
          
    'Descobre o address do range
    rangeAddress = Range(Cells(linhaInicial, colunaInicial), Cells(linhaFinal, colunaFinal)).Address

    'Apaga os dados anteriores na instancia
    Call ApagaDados(xls, sheetName, rangeAddress)
    
    'Seleciona o range que será colado os dados
    Me.workbookInstanciado.Sheets(sheetName).Range(rangeAddress).Value = dados
    
    '---------------------------------------------------
    '---------------------------------------------------
'    'Procura o workbook em alguma instancia do Excel que estiver aberta (Se estiver aberta, FOCA)
'    FocaWorkbook (pathBase)
'
'    'Se o arquivo ATUAL estiver aberto, cola nele mesmo
'    If ActiveWorkbook.FullName = pathBase Then
'
'        'Apaga os dados anteriores antes de colar os dados
'        If apagarDadosAnteriores = True Then
'
'            'Apaga os dados anteriores na própria Sheet
'            ApagaDadosSheet (sheetName)
'
'        End If
'
'        'Seleciona o range que será colado os dados
'        Sheets(sheetName).Range(rangeAddress).Value = dados
'
'        'Altera o data source da pivot table
'        Call ChangePivotSource(sheetName)
'
'    '---------------------------------------------------
'    '---------------------------------------------------
'    'Se não for o próprio arquivo, ABRE A INSTANCIA e cola os dados
'    Else
'
'        'Apaga os dados anteriores antes de colar os dados
'        If apagarDadosAnteriores = True Then
'
'            'Apaga os dados anteriores na instancia
'            Call ApagaDados(xls, sheetName, rangeAddress)
'
'        End If
'

'
'        'Altera o data source da pivot table
'        Call ChangePivotSourceInstancia(xls, sheetName)
'
'    End If
    
End Sub

'Procura o workbook em alguma instancia do Excel que estiver aberta (Se estiver aberta, FOCA)
Private Sub FocaWorkbook(ByVal pathBase As String)

    Dim Workbook As Workbook

    'Itera por todos os workbooks, se achar o path correto, foca nele na main window
    For Each Workbook In Workbooks
        
        'Se achar, foca nele
        If Workbook.FullName = pathBase Then
            Workbook.Windows(1).Activate
        End If
    Next Workbook
    
End Sub

'Altera o data source da pivot table
Private Sub ChangePivotSource(ByVal sheetName As String)
    
    Dim ultimaColuna As String: ultimaColuna = Sheets(sheetName).Range("A1").End(xlToRight).Address
    Dim linha As Long: linha = Sheets(sheetName).Range("A1").End(xlDown).Row
    
    'Se a última linha for a última da sheet o VALOR SE TORNA LINHA 2 (Atualizar a pivot)
    If linha = 1048576 Or _
       linha = 65536 Then
        linha = 2
    End If
    
    Dim ultimaLinha As String: ultimaLinha = Sheets(sheetName).Cells(linha, Sheets(sheetName).Range("A1").End(xlToRight).Column).Address
    Dim j As Long
    
    'Validando tamanho do array (Se for fora dos parametros, retorna erro de validação)
    If Sheets(sheetName).ProtectContents _
                            Or ultimaColuna = "$XFD$1" _
                            Or ultimaColuna = "$IV$1" _
                            Or ultimaLinha = "$XFD$1048576" Then
                
                Exit Sub
        
    End If
    
    'O endereço do range
    Dim rangeAddress As String: rangeAddress = "A1:" & ultimaLinha
    
    'declara o range da source da pivot table
    Dim rangeNovo As Range: Set rangeNovo = Sheets(sheetName).Range(rangeAddress)
    Dim pivotTable As pivotTable

    'Itera pelas pivots da sheet e muda o source da pivot
    For Each pivotTable In Sheets(sheetName).PivotTables
        
        pivotTable.ChangePivotCache ActiveWorkbook. _
            PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rangeNovo)
        
    Next pivotTable

End Sub

'Altera o data source da pivot table
Private Sub ChangePivotSourceInstancia(ByRef xls As Excel.Application, _
                                       ByVal sheetName As String)
                                       
    Dim ultimaColuna As String: ultimaColuna = Me.workbookInstanciado.Sheets(sheetName).Range("A1").End(xlToRight).Address
    Dim linha As Long: linha = Me.workbookInstanciado.Sheets(sheetName).Range("A1").End(xlDown).Row
    
    'Se a última linha for a última da sheet o VALOR SE TORNA LINHA 2 (Atualizar a pivot)
    If linha = 1048576 Or _
       linha = 65536 Then
        linha = 2
    End If
    
    Dim ultimaLinha As String: ultimaLinha = Me.workbookInstanciado.Sheets(sheetName).Cells(linha, Me.workbookInstanciado.Sheets(sheetName).Range("A1").End(xlToRight).Column).Address
    Dim j As Long
    
    'Validando tamanho do array (Se for fora dos parametros, retorna erro de validação)
    If Me.workbookInstanciado.Sheets(sheetName).ProtectContents _
                            Or ultimaColuna = "$XFD$1" _
                            Or ultimaColuna = "$IV$1" _
                            Or ultimaLinha = "$XFD$1048576" Then
                
                Exit Sub
        
    End If
    
    'O endereço do range que será
    Dim rangeAddress As String: rangeAddress = "A1:" & ultimaLinha
    
    'declara o range da source da pivot table
    Dim rangeNovo As Range: Set rangeNovo = Me.workbookInstanciado.Sheets(sheetName).Range(rangeAddress)
    Dim pivotTable As pivotTable
    
    
    'Desprotege a planilha
    Me.workbookInstanciado.Sheets(sheetName).Unprotect password:="dvd"
    
    'Itera pelas pivots da sheet e muda o source da pivot
    For Each pivotTable In Me.workbookInstanciado.Sheets(sheetName).PivotTables
        
        pivotTable.ChangePivotCache Me.workbookInstanciado. _
            PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rangeNovo)
        
    Next pivotTable
    
End Sub

'Apaga os dados anteriores na instancia
Private Sub ApagaDados(ByRef xls As Excel.Application, _
                       ByVal sheetName As String, _
                       ByVal rangeAcolar As String)

    Dim ultimaColuna As String
    ultimaColuna = Me.workbookInstanciado.Sheets(sheetName).Range("A1").End(xlToRight).Address
    
    Dim linha As Long
    linha = Me.workbookInstanciado.Sheets(sheetName).Range("A1").End(xlDown).Row
    
    Dim ultimaLinha As String
    ultimaLinha = Me.workbookInstanciado.Sheets(sheetName).Cells(linha, Me.workbookInstanciado.Sheets(sheetName).Range("A1").End(xlToRight).Column).Address
    
    Dim j As Long
    
    'Validando tamanho do array (Se for fora dos parametros, retorna erro de validação)
    If Me.workbookInstanciado.Sheets(sheetName).ProtectContents _
                            Or ultimaColuna = "$XFD$1" _
                            Or ultimaColuna = "$IV$1" _
                            Or ultimaLinha = "$XFD$1048576" Then
                
                Exit Sub
        
    End If
    
    'O endereço do range que será apagado
    Dim rangeAddress As String: rangeAddress = "A2:" & ultimaLinha
    
    Call ApagaDadosAnteriores(sheetName, rangeAddress, rangeAcolar)

End Sub

'Apaga os dados anteriores
Private Sub ApagaDadosAnteriores(ByVal sheetName As String, _
                                 ByVal rangeAddress As String, _
                                 ByVal rangeAcolar As String)
On Error GoTo BugFix

TryAgain:
    'Remove qualquer resíduo de filtro anterior
    Me.workbookInstanciado.Sheets(sheetName).Cells.AutoFilter
    
    'Apaga os dados anteriores
    Me.workbookInstanciado.Sheets(sheetName).Range(rangeAddress).ClearFormats
    Me.workbookInstanciado.Sheets(sheetName).Range(rangeAddress).ClearContents

    'Verifica se alguma Pivot está dentro do range que será apagado
    'Se houver muda o local da pivot para uma coluna a direita do range a ser apagado
    
    Exit Sub
BugFix:
    Call MudaPosicaoPivot(sheetName, rangeAcolar)
    GoTo TryAgain
    
End Sub

'Muda a posição da pivot que estiver dentro do range de dados
Private Sub MudaPosicaoPivot(ByVal sheetName As String, _
                             ByVal rangeAcolar As String)
    
    Dim pivotTable As pivotTable
    
    'Itera por todas as pivots da sheet
    For Each pivotTable In Me.workbookInstanciado.Sheets(sheetName).PivotTables
                
        'Verifica se a pivot está no range e se estiver, muda a posição
        Call MudaPivotRange(pivotTable, sheetName, rangeAcolar)
    
    Next pivotTable

End Sub

'Verifica se a pivot está no range e se estiver, muda a posição
Private Sub MudaPivotRange(ByRef pivotTable As pivotTable, _
                           ByVal sheetName As String, _
                           ByVal rangeAcolar As String)
    
    
    'Identifica a última coluna do address
    Dim ultimaColuna As Integer
    ultimaColuna = Range(rangeAcolar).Columns.Count
    
    Dim linhaDaPivot As Integer
    linhaDaPivot = pivotTable.RowRange.Row
    
    Dim novaPosicao As String
    novaPosicao = Cells(linhaDaPivot, ultimaColuna + 2).Address
    
    'Muda a posição da pivot, SEMPRE
    
    On Error GoTo AspaUnica
    pivotTable.Location = "'" & sheetName & "'!" & novaPosicao
    
    Exit Sub
AspaUnica:
    pivotTable.Location = novaPosicao
    'Me.workbookInstanciado.Sheets(sheetName).PivotTables("Tabela dinâmica2").Location = "' Analítico TFN'!$P$13"
    
End Sub

'Apaga os dados anteriores NA PRÓPRIA SHEET
Public Sub ApagaDadosSheet(ByVal sheetName As String)

    Dim ultimaColuna As String: ultimaColuna = Sheets(sheetName).Range("A1").End(xlToRight).Address
    Dim linha As Long: linha = Sheets(sheetName).Range("A1").End(xlDown).Row
    Dim ultimaLinha As String: ultimaLinha = Sheets(sheetName).Cells(linha, Sheets(sheetName).Range("A1").End(xlToRight).Column).Address
    Dim j As Long
    
    'Validando tamanho do array (Se for fora dos parametros, retorna erro de validação)
    If Sheets(sheetName).ProtectContents _
                                Or ultimaColuna = "$XFD$1" _
                                Or ultimaColuna = "$IV$1" _
                                Or ultimaLinha = "$XFD$1048576" Then
                
                Exit Sub
        
    End If
    
    'O endereço do range que será apagado
    Dim rangeAddress As String: rangeAddress = "A2:" & ultimaLinha
    
    'Seleciona o range de dados
    Sheets(sheetName).Range(rangeAddress).ClearContents
    
    'Altera a source da Pivot
    ChangePivotSource (sheetName)

End Sub

'Apaga os dados anteriores da RQ
Public Sub ApagaDadosDaRQ(ByVal sheetName As String, _
                          ByVal pathBase As String, _
                          ByVal apagarDadosAnteriores As Boolean)
    
    '---------------------------------------------------
    '---------------------------------------------------
    
    'Se não haver workbooks abertos, aí adiciona um
    If Workbooks.Count = 0 Then
        Workbooks.Add
    End If
    
    'Procura o workbook em alguma instancia do Excel que estiver aberta (Se estiver aberta, FOCA)
    FocaWorkbook (pathBase)
    
    'Se o arquivo ATUAL estiver aberto, cola nele mesmo
    If ActiveWorkbook.FullName = pathBase Then

        'Apaga os dados anteriores antes de colar os dados
        If apagarDadosAnteriores = True Then
            
            'Apaga os dados anteriores na própria Sheet
            ApagaDadosSheet (sheetName)
                        
        End If
        
        'Altera o data source da pivot table
        Call ChangePivotSource(sheetName)
    
    '---------------------------------------------------
    '---------------------------------------------------
    'Se não for o próprio arquivo, ABRE A INSTANCIA e apaga os dados
    Else
    
        'Apaga os dados anteriores antes de colar os dados
        If apagarDadosAnteriores = True Then
            
            'Apaga os dados anteriores na instancia
            'Call ApagaDados(xls, sheetName, )
                                   
        End If
                 
        'Altera o data source da pivot table
        Call ChangePivotSourceInstancia(xls, sheetName)
        
    End If

End Sub

'Salva a instancia na maquina
Public Sub Salva(ByVal pathBase As String)
    

    workbookInstanciado.Save

End Sub

'Valida se a header da instancia de RQ possui os mesmos parametros
Public Function HeaderError(ByVal sheetAondeCola As String, _
                            ByVal headerArray As Variant) As Boolean

    Dim ultimaColuna As String: ultimaColuna = Me.workbookInstanciado.Sheets(sheetAondeCola).Range("A1").End(xlToRight).Address
    Dim linha As Long: linha = Me.workbookInstanciado.Sheets(sheetAondeCola).Range("A1").End(xlDown).Row
    Dim ultimaLinha As String: ultimaLinha = Me.workbookInstanciado.Sheets(sheetAondeCola).Cells(linha, Me.workbookInstanciado.Sheets(sheetAondeCola).Range("A1").End(xlToRight).Column).Address
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
Public Sub CopiaNames(nomes() As name)

    Dim name As name
    Dim i As Integer: i = UBound(nomes)
    
    'Itera por todos os nomes do workbook e cria um array de nomes
    For Each name In Me.workbookInstanciado.Names
        
        'Redimensiona o array de acordo com o array de nomes
        i = i + 1
        ReDim Preserve nomes(1 To i)
        Set nomes(i) = name
        
    Next name

End Sub

'Captura os dados da Header
Public Function MontaHeaderArray(ByVal sheetDaBaseDeDados As String) As Variant
        
    Dim ultimaColuna As String: ultimaColuna = Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range("A1").End(xlToRight).Address
    Dim linha As Long: linha = Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range("A1").End(xlDown).Row
    Dim ultimaLinha As String: ultimaLinha = Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Cells(linha, Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range("A1").End(xlToRight).Column).Address
    Dim j As Long
    
    'Validando tamanho do array (Se for fora dos parametros, retorna erro de validação)
    If Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).ProtectContents _
                            Or ultimaColuna = "$XFD$1" _
                            Or ultimaColuna = "$IV$1" _
                            Or ultimaLinha = "$XFD$1048576" Then
                
                Exit Function
        
    End If
    
    'O endereço do range que será utilizado
    Dim rangeAddress As String: rangeAddress = "A1:" & ultimaColuna
    
    Dim i As Integer
    
    'Instancia só a header de sheetDaBaseDeDados
    Dim dataRangeHeader As Variant: dataRangeHeader = Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range(rangeAddress).Value
    
    'Retorna o array da header
    MontaHeaderArray = dataRangeHeader
                    
End Function

'Retorna a lista de sheets do arquivo
Public Function SheetsDoArquivo() As Variant

    Dim sheet As Worksheet
    Dim lista() As Variant
    Dim i As Integer
    
    'Itera pela lista de sheets
    For Each sheet In Me.workbookInstanciado.Worksheets
        
        i = i + 1
        ReDim Preserve lista(1 To i)
        
        'Adiciona cada sheet na listbox
        If Me.workbookInstanciado.ActiveSheet.name = sheet.name Then
            lista(i) = "*" & sheet.name
        Else
            lista(i) = sheet.name
        End If
        
    Next sheet
    
    'Retorna a lista de sheets do arquivo
    SheetsDoArquivo = lista

End Function

'Desprotege todas as planilhas do workbook
Public Sub DesprotegePlanilhas()
    
    Dim sheet As Worksheet
    
    'Itera por todas as sheets do workbook desprotegendo
    For Each sheet In Me.workbookInstanciado.Sheets
        
    Next sheet
    
End Sub

'Captura os dados das colunas
Public Function CapturaColunas(ByRef sheetName As String) As Variant
    
    Dim ultimaColuna As String: ultimaColuna = Me.workbookInstanciado.Sheets(sheetName).Range("A1").End(xlToRight).Address
    Dim linha As Long: linha = Me.workbookInstanciado.Sheets(sheetName).Range("A1").End(xlDown).Row
    Dim ultimaLinha As String: ultimaLinha = Me.workbookInstanciado.Sheets(sheetName).Cells(linha, Me.workbookInstanciado.Sheets(sheetName).Range("A1").End(xlToRight).Column).Address
    Dim j As Long
    
    'Validando tamanho do array (Se for fora dos parametros, retorna erro de validação)
    If ultimaColuna = "$XFD$1" _
        Or ultimaColuna = "$IV$1" _
        Or ultimaLinha = "$XFD$1048576" Then
                
            Exit Function
    End If
    
    'O endereço do range que será capturado
    Dim rangeAddress As String: rangeAddress = "A1:" & ultimaColuna
    Dim dataRange As Variant

    Me.workbookInstanciado.Sheets(sheetName).Range(rangeAddress).NumberFormat = "General" 'Corrigindo bug
    
    'Array de dados
    dataRange = Me.workbookInstanciado.Sheets(sheetName).Range(rangeAddress).Value
    
    'Retorna os dados do array (Transpõe cada coluna em linha)
    CapturaColunas = Application.WorksheetFunction.Transpose(dataRange)

End Function

'Protege todas as planilhas do workbook
Public Sub ProtegePlanilhas()

    Dim sheet As Worksheet
    
    'Itera por todas as sheets do workbook desprotegendo
    For Each sheet In Me.workbookInstanciado.Sheets
        
    Next sheet

End Sub

'Monta um array de 2 colunas contendo os dados de valores únicos sobre cada valor da coluna
Public Function MontaValoresUnicos(ByRef listaDeColunas As Variant, _
                                   ByVal sheetName As String) As Variant
        
    Dim i As Long
    Dim listaValoresUnicos As Variant
    Dim listaCompleta() As Variant
    Dim j As Integer
    Dim ItemsUnicos As ItemsUnicos: Set ItemsUnicos = New ItemsUnicos
    
    'Se houver mais de 15 colunas, exibir msg de erro
'    If UBound(listaDeColunas) > 40 Then
'        MsgBox "A quantidade de colunas excede ao limite de 40 colunas."
'        Exit Function
'    End If
    
    'Se não houver dados na lista de colunas
    If IsArray(listaDeColunas) = False Then
        Exit Function
    End If

    'lista de valores únicos de cada coluna
    For i = 1 To UBound(listaDeColunas)
        
        'Pega os valores únicos de cada coluna
        listaValoresUnicos = ValoresUnicosDic(i, sheetName)
        
        'Adiciona os dados na coleção
        Call AddCol(ItemsUnicos, listaValoresUnicos)
    Next i
    
    'Transforma a coleção em array
    MontaValoresUnicos = ColectionToArray(ItemsUnicos)

End Function

'Transforma a coleção em array
Private Function ColectionToArray(ByRef ItemsUnicos As ItemsUnicos) As Variant
    
    'Tamanho do array
    Dim tamanhoDoArray As Long
    tamanhoDoArray = ItemsUnicos.Count
    
    'Redimensiona o array
    Dim lista As Variant
    ReDim lista(1 To tamanhoDoArray, 1 To 2)
    
    Dim i As Long
    Dim ItemUnico As ItemUnico
    
    'Itera pela coleção e monta o array
    For Each ItemUnico In ItemsUnicos
        
        i = i + 1
        lista(i, 1) = ItemUnico.coluna
        lista(i, 2) = ItemUnico.valor
    
    Next ItemUnico
    
    'Retorna a lista com os valores
    ColectionToArray = lista
    
End Function

'Adiciona os dados na coleção
Public Sub AddCol(ByRef ItemsUnicos As ItemsUnicos, _
                  ByRef listaValoresUnicos As Variant)
    
    Dim i As Integer
    
    Dim ItemUnico As ItemUnico
    
    'Adiciona os valores únicos em cada coluna
    For i = 1 To UBound(listaValoresUnicos)
        
        Set ItemUnico = New ItemUnico
        
        'Cria os valores da coleção
        ItemUnico.coluna = listaValoresUnicos(0)
        ItemUnico.valor = listaValoresUnicos(i)
        
        'Pra cada linha do array, adiciona os valores na coleção
        ItemsUnicos.Add ItemUnico
    Next i

End Sub

'Pega os valores únicos de cada coluna
Private Function ValoresUnicosDic(ByVal colunaNumero As Integer, _
                                  ByVal sheetName As String) As Variant

    'Cria um novo dicionário
    Dim dic As Scripting.Dictionary: Set dic = New Dictionary
    Dim i As Long

    'Monta o dicionário com os valores únicos
    Dim ultimaLinha As Long
    ultimaLinha = Me.workbookInstanciado.Sheets(sheetName).UsedRange.Rows.Count
    
    Dim cellValue As String
    
    Dim endereco As String
    endereco = Range(Cells(1, colunaNumero), Cells(ultimaLinha, colunaNumero)).Address
    
    Dim dataRange As Variant
    dataRange = Me.workbookInstanciado.Sheets(sheetName).Range(endereco).Value
    
    'Itera pela lista da sheet e monta os valores únicos
    For i = 1 To ultimaLinha
            
        'Valor da célula
        cellValue = dataRange(i, 1)
            
        'Coloca no dicionário apenas os registros únicos
        If dic.Exists(cellValue) = False Then
            
            dic.Add cellValue, cellValue
        End If
    Next i
    
    'Itera pelos valores do dicionário e retorna a lista preenchida
    ValoresUnicosDic = dic.Items

End Function

'Cria a sheet no instancia
Sub CriaSheet(ByRef sheetAondeCola As String)
    
On Error GoTo exitSub
        
    Dim qtdSheets As Integer
    qtdSheets = Me.workbookInstanciado.Sheets.Count
    
    'Adiciona a sheet como primeira sheet
    Me.workbookInstanciado.Sheets.Add After:=Me.workbookInstanciado.Sheets(qtdSheets)
    
    'Renomeia a sheet adicionada
    Me.workbookInstanciado.Sheets(qtdSheets + 1).name = sheetAondeCola

exitSub:
End Sub

'Filtra os dados de acordo com os filtros
Public Function FiltraDadosDataRange(ByVal filtros As Variant, _
                                     ByVal sheetDaBaseDeDados As String) As Variant
    
    Dim tamanhoDoArray As Long
    Dim novoDataRange() As Variant
    Dim linhaNovoDataRange As Integer
    Dim dataRange As Variant
    
    'Filtra os dados do datarange e captura
    dataRange = CapturaDadosDataRange(filtros, sheetDaBaseDeDados)
    
    'Valida se o array é vazio
    If IsEmpty(dataRange) Then
        Exit Function
    End If
    
    'Retorna o array de dados filtrados
    FiltraDadosDataRange = dataRange
    
End Function

'Filtra os dados do datarange e captura
Private Function CapturaDadosDataRange(ByVal filtros As Variant, _
                                       ByVal sheetDaBaseDeDados As String) As Variant
        
    'Remove qualquer resíduo de filtro anterior
    Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Cells.AutoFilter
    
    'Dados da planilha
    '----------------------------------
    Dim ultimaColuna As String: ultimaColuna = Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range("A1").End(xlToRight).Address
    Dim linha As Long: linha = Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range("A1").End(xlDown).Row
    Dim ultimaLinha As String: ultimaLinha = Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Cells(linha, Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range("A1").End(xlToRight).Column).Address
    Dim j As Long
    
    'Validando tamanho do array (Se for fora dos parametros, retorna erro de validação)
    If ultimaColuna = "$XFD$1" _
      Or ultimaColuna = "$IV$1" _
      Or ultimaLinha = "$XFD$1048576" Then
                
                Exit Function
    End If
    
'*************************************
'****** FILTROS
    
    'Identifica a posição da coluna
    Dim coluna As Integer
    Dim ouArray As Variant
    Dim i As Integer
    Dim valorCriterios As Variant
    Dim quotes As String
    quotes = """"
    
    Dim rangeFiltro As String, dataRangeFiltro As Variant
    rangeFiltro = "A1:" & ultimaColuna
    dataRangeFiltro = Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range(rangeFiltro).Value
    
    'Insere os filtros no range
    Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range(rangeFiltro).AutoFilter
    
    'Se não houver filtros, vai pra próxima pega o range inteiro
    If IsArray(filtros) Then
    
        'Itera pelo cabeçalho(filtros) em busca do filtro *posição numérica
        For j = 1 To UBound(filtros)
            
            'Iteração de cabeçalho
            For i = 1 To UBound(dataRangeFiltro, 2)
                
                'Se o valor do filtro for igual ao do cabeçalho
                If filtros(j, 1) Like dataRangeFiltro(1, i) Then
                
                    'Se for clausula "E"
                    If InStr(1, filtros(j, 2), ",") = 0 Then
                        'Quando encontra o filtro, insere no cabeçalho
                        Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range(rangeFiltro).AutoFilter _
                                                field:=i, _
                                                Criteria1:=Replace(filtros(j, 2), quotes, "")
                    'Se for clausula "OU"
                    Else
                        
                        'Se há mais de 3 argumentos no array
                        If UBound(ConverteEmArray(filtros(j, 2))) > 2 Then
                            'Marca os filtros
                            Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range(rangeFiltro).AutoFilter _
                                                        field:=i, _
                                                        Criteria1:=ConverteEmArray(filtros(j, 2)), _
                                                        Operator:=xlFilterValues
                        'Se For igual a 2 criterios
                        Else
                            valorCriterios = ConverteEmArray(filtros(j, 2))
                            Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range(rangeFiltro).AutoFilter _
                                    field:=i, _
                                    Criteria1:="=" & valorCriterios(0), _
                                    Operator:=xlOr, _
                                    Criteria2:="=" & valorCriterios(1)
                        End If
                        
                    End If
                End If
            Next i
        Next j
    End If
    
    Dim rangeAddress As String
    rangeAddress = "A1" & ":" & ultimaLinha
    
'----------------------------------------------------------------
    'Copia o datarange
    Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range(rangeAddress).Copy
    
    Application.Wait Now + TimeValue("00:00:05")
    
    'Apaga os dados anteriores em paste e cola na sheet do programa para usar os dados
    meWBK.Sheets("_Paste").UsedRange.Clear
    meWBK.Sheets("_Paste").Visible = xlSheetVisible
    

'Encapsula o copy and paste
    Call CopyAndPaste(xls, sheetDaBaseDeDados, rangeAddress)

    'Seleciona o range a copiar
    CapturaDadosDataRange = Sheets("_Paste").UsedRange.Value
        
End Function

'Encapsula o copy and paste
Private Sub CopyAndPaste(ByRef xls As Excel.Application, _
                         ByVal sheetDaBaseDeDados As String, _
                         ByVal rangeAddress As String)

    On Error GoTo TryAgain
    
    Me.workbookInstanciado.Sheets(sheetDaBaseDeDados).Range(rangeAddress).Copy
    Application.Wait Now + TimeValue("00:00:01")
    Excel.ActiveWorkbook.Sheets("_Paste").Paste
    
Exit Sub

TryAgain:
    Call CopyAndPaste(xls, sheetDaBaseDeDados, rangeAddress)
End Sub

'Converte em array o "OU"
Private Function ConverteEmArray(ByVal filtros As String) As Variant
        
    'conta a quantidade de vírgulas
    Dim i As Integer
    Dim contador As Integer
    
    'Itera pelas letras de filtros
    For i = 1 To Len(filtros)
        
        'Quantida de vírgulas
        If Mid(filtros, i, 1) = "," Then
            contador = contador + 1
        End If
    Next i
    
    'monta o array de filtros (quantidade de vírgulas +1)
    Dim arrayFiltros As Variant
    arrayFiltros = Split(filtros, ",")
    
    'Substitui os caracteres em branco do array (Iteração)
    Dim j As Integer
    Dim quote As String
    Dim palavraProcurada As String
    
    quote = """"
    palavraProcurada = " " & quote
    
    'Itera pelos valores do array
    For j = LBound(arrayFiltros) To UBound(arrayFiltros)
        arrayFiltros(j) = Replace(arrayFiltros(j), palavraProcurada, quote)
        arrayFiltros(j) = Replace(arrayFiltros(j), quote, "")
    Next j
    
    'Retorna o array montado
    ConverteEmArray = arrayFiltros
    
End Function
