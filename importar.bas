Attribute VB_Name = "importar"
Option Compare Database
Option Explicit

'Seleciona o Excel que sera importado
Public Function SelecionaExcel(ByRef arquivo_endereco As TextBox) As String
  
  Dim FileDialogs As FileDialogs
  Set FileDialogs = New FileDialogs
  
  arquivo_endereco.Value = FileDialogs.SelectFile("Selecione o arquivo")
  
End Function

'Importa os dados selecionados
Public Sub importarDados(ByRef frm_importar As Form, _
                         sheetName As String, _
                         linhaInicial As Integer, _
                         colunaInicial As Integer, _
                         colunaFinal As Integer, _
                         colunaInicialNumerica As Integer, _
                         ByRef msgsLog As Variant, _
                         validacaoNumerica As Boolean)
  Dim Load As Load
  Set Load = New Load
  Dim CriandoObjetos As CriandoObjetos
  Set CriandoObjetos = New CriandoObjetos
  
  Dim InstanciaExcel As InstanciaExcel
  Set InstanciaExcel = New InstanciaExcel
  
  Dim k As Integer
  Dim importInicio As Double
  importInicio = Timer
  '---------------------------------------------------------------------
  'verifica se o tipo de relatório foi selecionado e o arquivo
  If ((frm_importar.reportTipos.Value = "" Or IsNull(frm_importar.reportTipos.Value)) Or _
     (frm_importar.arquivo_endereco.Value = "" Or IsNull(frm_importar.arquivo_endereco.Value)) Or _
     (frm_importar.dataMesAno.Value = "" Or IsNull(frm_importar.dataMesAno.Value))) Then
    MsgBox "É necessário selecionar o 'tipo' de relatório e o 'arquivo' que serão importados.", vbCritical, "Selecione os campos"
    Exit Sub
  End If
  
  '---------------------------------------------------------------------
  'verifica se o arquivo existe
  k = k + 1
  msgsLog(k, 1) = "Identifica se o arquivo é válido"
  If CriandoObjetos.ArquivoExiste(frm_importar.arquivo_endereco.Value) = False Then
      msgsLog(k, 2) = "NOK - Selecione um arquivo válido"
    Exit Sub
  End If
  msgsLog(k, 2) = "OK"
    
  'se o arquivo existe instancia ele
  Call InstanciaExcel.InstanciaObjeto(frm_importar.arquivo_endereco.Value)
  
  '---------------------------------------------------------------------
  'verifica se a sheet procurada existe
  k = k + 1
  msgsLog(k, 1) = "Procura pela Sheet '" & sheetName & "'"
  If InstanciaExcel.ValidaSeSheetExiste(sheetName) = False Then
    msgsLog(k, 2) = "NOK - Sheet '" & sheetName & "' não encontrada"
    Exit Sub
  End If
  msgsLog(k, 2) = "OK"
  
  '---------------------------------------------------------------------
  'Validacao de dados pelo cabecalho (CABECALHO)
  Dim cabecalho As Variant
  cabecalho = defineCabecalho(frm_importar.reportTipos.Value)
  Dim datarangeCabecalho As Variant
  datarangeCabecalho = InstanciaExcel.xls.Sheets(sheetName).Range(InstanciaExcel.xls.Sheets(sheetName).Cells(linhaInicial, colunaInicial), _
                                                                  InstanciaExcel.xls.Sheets(sheetName).Cells(linhaInicial, colunaFinal)).Value
  Dim ultimaLinha As Double
  ultimaLinha = InstanciaExcel.xls.Sheets(sheetName).Cells(InstanciaExcel.xls.Rows.Count, colunaFinal).End(xlUp).Row
    
  'verifica se o nome das colunas batem
  k = k + 1
  msgsLog(k, 1) = "Válida cabeçalho de dados"
  If validaCabecalhoNomes(datarangeCabecalho, cabecalho) = True Then
    msgsLog(k, 2) = "Os dados do cabeçalho estão diferentes da parametrização interna"
    Exit Sub
  End If
  msgsLog(k, 2) = "OK"
  
  'verifica se existem dados na sheet
  k = k + 1
  msgsLog(k, 1) = "Arquivo com dados"
  If InstanciaExcel.existeDadosSheet(sheetName, linhaInicial, colunaInicial) = False Then
    msgsLog(k, 2) = "Não existem dados na planilha"
    Exit Sub
  End If
  msgsLog(k, 2) = "OK"
  
  'todos os dados =
  Dim DataRange As Variant
  DataRange = InstanciaExcel.xls.Sheets(sheetName).Range(InstanciaExcel.xls.Sheets(sheetName).Cells(linhaInicial, colunaInicial), _
                                                         InstanciaExcel.xls.Sheets(sheetName).Cells(ultimaLinha, colunaFinal)).FormulaLocal
  'so os dados numericos
  Dim datarangeNumerico As Variant
  datarangeNumerico = InstanciaExcel.xls.Sheets(sheetName).Range(InstanciaExcel.xls.Sheets(sheetName).Cells(linhaInicial + 1, colunaInicialNumerica), _
                                                                 InstanciaExcel.xls.Sheets(sheetName).Cells(ultimaLinha, colunaFinal)).FormulaLocal
                                                                 
  'valida o tipo de dados numericos na sheet e se há formulas com erros
  k = k + 1
  msgsLog(k, 1) = "Erros de fórmula?"
  If validaTiposDeDados(DataRange, datarangeNumerico, cabecalho, sheetName, ultimaLinha, validacaoNumerica) = False Then
    msgsLog(k, 2) = "Existem erros de fórmulas na planilha"
    Exit Sub
  End If
  msgsLog(k, 2) = "OK"
    
  '---------------------------------------------------------------------
  'Cola os dados na tabela temp
  'da o "INSERT" na tabela "input_monthBase"
  'quando importar os dados com sucesso
  k = k + 1
  msgsLog(k, 1) = "DADOS IMPORTADOS"
  If cargaImportacao(frm_importar.reportTipos.Value, DataRange, Form_frm_importar.dataMesAno.Value) = True Then
    msgsLog(k, 2) = "Erro na carga de dados"
    Exit Sub
  End If
  msgsLog(k, 2) = "***** OK: " & frm_importar.reportTipos.Value & " | importação: " & Round(Timer - importInicio, 2) & "s"
  
End Sub

'da a carga de dados na base de dados
Public Function cargaImportacao(nomeReport As String, _
                                DataRange As Variant, _
                                dataMesAno As Date) As Boolean
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Dim i As Long
  Dim sSQL As String
  
  Dim COnverter As COnverter
  Set COnverter = New COnverter
  
'On Error GoTo trataErro

  '++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  If nomeReport = "Input Month Base" Then

    'apaga os dados da temp
    DoCmd.SetWarnings False
    DoCmd.RunSQL "delete * from input_monthBase"
    DoCmd.RunSQL "delete * from temp_input_monthBase"
    DoCmd.SetWarnings True

    'da a carga na temp '(Insere novos dados aqui se a validação de colunas estiver ok)
    Set db = DBEngine.OpenDatabase(CurrentProject.FullName)
    Set rs = db.OpenRecordset("select * from temp_input_monthBase", dbOpenDynaset)

    For i = 2 To UBound(DataRange)
        rs.AddNew
        rs.Fields("data").Value = dataMesAno
        rs.Fields("SKU").Value = DataRange(i, 1)
        rs.Fields("SKU Description").Value = DataRange(i, 2)
        rs.Fields("Category").Value = DataRange(i, 3)
        rs.Fields("Regions_Channel").Value = DataRange(i, 4)
        rs.Fields("BU").Value = DataRange(i, 5)
        rs.Fields("Brand").Value = DataRange(i, 6)
        rs.Fields("actual_tons").Value = COnverter.importaNumeros(DataRange(i, 8))
        rs.Fields("actual_lsv").Value = COnverter.importaNumeros(DataRange(i, 9))
        rs.Fields("actual_nsv").Value = COnverter.importaNumeros(DataRange(i, 10))
        rs.Fields("actual_vic").Value = COnverter.importaNumeros(DataRange(i, 11))
        rs.Fields("actual_vlc").Value = COnverter.importaNumeros(DataRange(i, 12))
        rs.Fields("budget_tons").Value = COnverter.importaNumeros(DataRange(i, 14))
        rs.Fields("budget_lsv").Value = COnverter.importaNumeros(DataRange(i, 15))
        rs.Fields("budget_nsv").Value = COnverter.importaNumeros(DataRange(i, 16))
        rs.Fields("budget_vic").Value = COnverter.importaNumeros(DataRange(i, 17))
        rs.Fields("budget_vlc").Value = COnverter.importaNumeros(DataRange(i, 18))
        rs.Fields("py_tons").Value = COnverter.importaNumeros(DataRange(i, 20))
        rs.Fields("py_lsv").Value = COnverter.importaNumeros(DataRange(i, 21))
        rs.Fields("py_nsv").Value = COnverter.importaNumeros(DataRange(i, 22))
        rs.Fields("py_vic").Value = COnverter.importaNumeros(DataRange(i, 23))
        rs.Fields("py_vlc").Value = COnverter.importaNumeros(DataRange(i, 24))
        rs.Update
    Next i
    Set rs = Nothing
    
    '-----------------------------------------------
    'da temp da a carga na tabela
    sSQL = "insert into input_monthBase " & _
           " ( data, SKU, [SKU Description], Category, Regions_Channel, BU, Brand, actual_tons, actual_lsv, actual_nsv, actual_vic, actual_vlc, budget_tons, budget_lsv, budget_nsv, budget_vic, budget_vlc, py_tons, py_lsv, py_nsv, py_vic, py_vlc ) " & _
           "select " & _
           "   data, SKU, [SKU Description], Category, Regions_Channel, BU, Brand, Sum(actual_tons) AS _actual_tons, Sum(actual_lsv) AS _actual_lsv, Sum(actual_nsv) AS _actual_nsv, Sum(actual_vic) AS _actual_vic, Sum(actual_vlc) AS _actual_vlc, Sum(budget_tons) AS _budget_tons, Sum(budget_lsv) AS _budget_lsv, Sum(budget_nsv) AS _budget_nsv, Sum(budget_vic) AS _budget_vic, Sum(budget_vlc) AS _budget_vlc, Sum(py_tons) AS _py_tons, Sum(py_lsv) AS _py_lsv, Sum(py_nsv) AS _py_nsv, Sum(py_vic) AS _py_vic, Sum(py_vlc) AS _py_vlc " & _
           "from " & _
           " temp_input_monthBase " & _
           "group by " & _
           " data, SKU, [SKU Description], Category, Regions_Channel, BU, Brand;"
                      
    DoCmd.SetWarnings False
    DoCmd.RunSQL sSQL
       
    '-----------------------------------------------------
    'da a carga de dados na tabela 'base_report_INFERIOR_t' e na base de dados
    
    Dim rsBU As Recordset
    Set rsBU = CurrentDb.OpenRecordset("select distinct BU from temp_input_monthBase")
    
    'itera pelas tabelas
    While Not rsBU.EOF
      DoCmd.RunSQL "delete * from base_report_INFERIOR_t where data=#" & Format(dataMesAno, "mm/dd/yyyy") & "# and BU= '" & rsBU.Fields("BU").Value & "'"
      
      DoCmd.RunSQL "delete * from base_report_INFERIOR_t_bd where data=#" & Format(dataMesAno, "mm/dd/yyyy") & "# and BU= '" & rsBU.Fields("BU").Value & "'"
      DoCmd.RunSQL "delete * from categoryMIX_report_final_AOP_adju_bd where data=#" & Format(dataMesAno, "mm/dd/yyyy") & "# and BU= '" & rsBU.Fields("BU").Value & "'"
      DoCmd.RunSQL "delete * from categoryMIX_report_final_AOP_bd where data=#" & Format(dataMesAno, "mm/dd/yyyy") & "# and BU= '" & rsBU.Fields("BU").Value & "'"
      DoCmd.RunSQL "delete * from categoryMIX_report_PY_bd where data=#" & Format(dataMesAno, "mm/dd/yyyy") & "# and BU= '" & rsBU.Fields("BU").Value & "'"
      DoCmd.RunSQL "delete * from MIX_total_Categoria_Final_AOP_bd where data=#" & Format(dataMesAno, "mm/dd/yyyy") & "# and BU= '" & rsBU.Fields("BU").Value & "'"
      DoCmd.RunSQL "delete * from MIX_total_Categoria_Final_PY_II_final_bd where data=#" & Format(dataMesAno, "mm/dd/yyyy") & "# and BU= '" & rsBU.Fields("BU").Value & "'"
      DoCmd.RunSQL "delete * from categoryMIX_report_PY_adj_bd where data=#" & Format(dataMesAno, "mm/dd/yyyy") & "# and BU= '" & rsBU.Fields("BU").Value & "'"
      DoCmd.RunSQL "delete * from regionMIX_report_bd where data=#" & Format(dataMesAno, "mm/dd/yyyy") & "# and BU= '" & rsBU.Fields("BU").Value & "'"
      DoCmd.RunSQL "delete * from regionMIX_report_PY_bd where data=#" & Format(dataMesAno, "mm/dd/yyyy") & "# and BU= '" & rsBU.Fields("BU").Value & "'"
      
      rsBU.MoveNext
    Wend
       
    Call cargaEmBASE_report_INFERIOR
    
    '-----------------------------------------------------
    'cria registros com categorias com adj e sem adju ****** IMPORTANTE ******
    Call cargaCategoriasFaltantes
    
    '-----------------------------------------------------
    'Call updateSKUMIX_PY 'atualiza os dados de SKUMIX do relatório de PY
    '-----------------------------------------------------
    
    DoCmd.RunSQL "delete * from temp_input_monthBase"
    DoCmd.RunSQL "delete * from input_monthBase"
    
    '-----------------------------------------------------
    'carrega bd - 5 tabelas de dados
    DoCmd.RunSQL "insert into base_report_INFERIOR_t_bd select * from base_report_INFERIOR_t"
    Call cargaNosBdsAgrupados(Format(dataMesAno, "mm/dd/yyyy")) 'da a carga nos bd agrupados (4 tabelas bd)
    
    'da a carga na tabela "categoriaOthers"'
    Call cargaCategoriaOthers

    DoCmd.RunSQL "delete * from base_report_INFERIOR_t"
    '-----------------------------------------------------
    
    DoCmd.SetWarnings True
  End If

Exit Function
'------------------------------------------------
trataErro:
  cargaImportacao = True
End Function

'da a carga na tabela "categoriaOthers"'
Public Sub cargaCategoriaOthers()
  
  Dim sSQL As String
  Dim deleteSQL As String
  Dim rsDelete As Recordset
  
  '----------------------------------------------------------------------------------------
  'excluir os itens de 'categoriaOthers' que nao estiverem em 'base_report_INFERIOR_t_bd'
              deleteSQL = " select categoriaOthers.*  "
  deleteSQL = deleteSQL & "     from categoriaOthers left join base_report_INFERIOR_t_bd on  "
  deleteSQL = deleteSQL & "       (base_report_INFERIOR_t_bd.BU = categoriaOthers.BU) AND   "
  deleteSQL = deleteSQL & "       (base_report_INFERIOR_t_bd.Category = categoriaOthers.Category) "
  deleteSQL = deleteSQL & "    "
  deleteSQL = deleteSQL & "   where  "
  deleteSQL = deleteSQL & "     base_report_INFERIOR_t_bd.BU is null and  "
  deleteSQL = deleteSQL & "     base_report_INFERIOR_t_bd.Category is null"
  
  Set rsDelete = CurrentDb.OpenRecordset(deleteSQL)
  
  'itera pelo recordset excluindo
  While Not rsDelete.EOF
    DoCmd.RunSQL "delete * from categoriaOthers where BU = '" & rsDelete.Fields("BU").Value & "' and Category = '" & rsDelete.Fields("Category").Value & "'"
    rsDelete.MoveNext
  Wend
  
  '-----------------------------------------------------
  'carga na tabela 'categoriaOthers'
  'incluir os itens de 'base_report_INFERIOR_t_bd' que nao estiverem em  que 'categoriaOthers'
        sSQL = " INSERT INTO  "
  sSQL = sSQL & "   categoriaOthers ( BU, Category) "
  sSQL = sSQL & " SELECT DISTINCT  "
  sSQL = sSQL & "   base_report_INFERIOR_t_bd.BU, base_report_INFERIOR_t_bd.Category "
  sSQL = sSQL & " FROM  "
  sSQL = sSQL & "   base_report_INFERIOR_t_bd LEFT JOIN categoriaOthers ON base_report_INFERIOR_t_bd.BU = categoriaOthers.BU AND base_report_INFERIOR_t_bd.Category = categoriaOthers.Category "
  sSQL = sSQL & " WHERE  "
  sSQL = sSQL & "   categoriaOthers.BU Is Null AND categoriaOthers.Category Is Null "

  DoCmd.RunSQL sSQL

End Sub


'da a carga nos bd agrupados (4 tabelas bd)
Public Sub cargaNosBdsAgrupados(dataMesAno As String)

  Dim inserTables(1 To 8, 1 To 2) As String
  Dim i As Integer
  Dim sSQL As String
  
  inserTables(1, 1) = "MIX_total_Categoria_Final_AOP"
  inserTables(1, 2) = "MIX_total_Categoria_Final_AOP_bd"
  
  inserTables(2, 1) = "categoryMIX_report_final_AOP"
  inserTables(2, 2) = "categoryMIX_report_final_AOP_bd"
  
  inserTables(3, 1) = "MIX_total_Categoria_Final_PY_II_final"
  inserTables(3, 2) = "MIX_total_Categoria_Final_PY_II_final_bd"
  
  inserTables(4, 1) = "categoryMIX_report_PY"
  inserTables(4, 2) = "categoryMIX_report_PY_bd"

  inserTables(5, 1) = "categoryMIX_report_final_AOP_adju"
  inserTables(5, 2) = "categoryMIX_report_final_AOP_adju_bd"
  
  inserTables(6, 1) = "categoryMIX_report_PY_adj"
  inserTables(6, 2) = "categoryMIX_report_PY_adj_bd"
  
  inserTables(7, 1) = "regionMIX_report"
  inserTables(7, 2) = "regionMIX_report_bd"
  
  inserTables(8, 1) = "regionMIX_report_PY"
  inserTables(8, 2) = "regionMIX_report_PY_bd"


  'da a carga nas tabelas e exclui o valor anterior
  For i = 1 To UBound(inserTables)
    'DoCmd.RunSQL "delete * from " & inserTables(i, 2) & " where data =#" & dataMesAno & "#"
    DoCmd.RunSQL "insert into " & inserTables(i, 2) & " select * from " & inserTables(i, 1)
  Next i

End Sub

'atualiza os dados de SKUMIX do relatório de PY
Public Sub updateSKUMIX_PY()

  Dim sSQL As String
    
  DoCmd.RunSQL "delete * from skuMIX_PY_temp" 'apaga os dados da temp
  DoCmd.RunSQL "insert into skuMIX_PY_temp select * from skuMIX_py_qry"

           sSQL = " UPDATE  "
  sSQL = sSQL & "   skuMIX_PY_temp INNER JOIN base_report_INFERIOR_t ON  "
  sSQL = sSQL & "   (skuMIX_PY_temp.Regions_Channel = base_report_INFERIOR_t.Regions_Channel) AND  "
  sSQL = sSQL & "   (skuMIX_PY_temp.Category = base_report_INFERIOR_t.Category) AND  "
  sSQL = sSQL & "   (skuMIX_PY_temp.SKU = base_report_INFERIOR_t.SKU) AND  "
  sSQL = sSQL & "   (skuMIX_PY_temp.data = base_report_INFERIOR_t.data)  "
  sSQL = sSQL & " SET  "
  sSQL = sSQL & "   base_report_INFERIOR_t.skuMIX_PY_VIC = [skuMIX_PY_VIC_U],  "
  sSQL = sSQL & "   base_report_INFERIOR_t.skuMIX_PY_VLC = [skuMIX_PY_VLC_U],  "
  sSQL = sSQL & "   base_report_INFERIOR_t.skuMIX_PY_CMA = [skuMIX_PY_CMA_U] "
  DoCmd.RunSQL sSQL 'run
  
  DoCmd.RunSQL "delete * from skuMIX_PY_temp" 'apaga os dados da temp
    
End Sub

'cria registros com categorias com adj e sem adju ****** IMPORTANTE ******
Sub cargaCategoriasFaltantes()
  
  Dim rs As Recordset
  Dim rs_base_report_inferior As Recordset
  Dim coluna As Field
  Dim sSQL As String
  
  '------------------------------------------------
  'com adj

           sSQL = " SELECT Categories.BU as BU_correta,  Categories.data as data_correta, Categories.Category AS cat, MIX_total_Categoria_comAdj.Category "
    sSQL = sSQL & " FROM Categories LEFT JOIN MIX_total_Categoria_comAdj ON (Categories.Category = MIX_total_Categoria_comAdj.Category) AND (Categories.Data = MIX_total_Categoria_comAdj.data) "
    sSQL = sSQL & " WHERE (((MIX_total_Categoria_comAdj.Category) Is Null)) "
  
  Set rs = CurrentDb.OpenRecordset(sSQL)
  Set rs_base_report_inferior = CurrentDb.OpenRecordset("select * from base_report_INFERIOR_t")
  
  While Not rs.EOF
    
    rs_base_report_inferior.AddNew 'cria um novo registro
    For Each coluna In rs_base_report_inferior.Fields
      coluna.Value = 0
    Next coluna
    rs_base_report_inferior.Fields("adj").Value = 0
    rs_base_report_inferior.Fields("BU").Value = rs.Fields("BU_correta").Value
    rs_base_report_inferior.Fields("Category").Value = rs.Fields("cat").Value
    rs_base_report_inferior.Fields("data").Value = rs.Fields("data_correta").Value
    
    rs_base_report_inferior.Update 'insere um novo registro
    
    rs.MoveNext
  Wend
  
  '------------------------------------------------
  'sem adj
         sSQL = " select Categories.BU as BU_correta, Categories.data as data_correta, Categories.Category as cat, MIX_total_Categoria_semAdj.Category "
  sSQL = sSQL & " from MIX_total_Categoria_semAdj right join Categories on MIX_total_Categoria_semAdj.Category = Categories.Category "
  sSQL = sSQL & " where "
  sSQL = sSQL & " MIX_total_Categoria_semAdj.Category is null "
  
  Set rs = CurrentDb.OpenRecordset(sSQL)
  Set rs_base_report_inferior = CurrentDb.OpenRecordset("select * from base_report_INFERIOR_t")
  
  While Not rs.EOF
    
    rs_base_report_inferior.AddNew 'cria um novo registro
    For Each coluna In rs_base_report_inferior.Fields
      coluna.Value = 0
    Next coluna
    rs_base_report_inferior.Fields("adj").Value = 1
    rs_base_report_inferior.Fields("BU").Value = rs.Fields("BU_correta").Value
    rs_base_report_inferior.Fields("Category").Value = rs.Fields("cat").Value
    rs_base_report_inferior.Fields("data").Value = rs.Fields("data_correta").Value
    rs_base_report_inferior.Update 'insere um novo registro
    
    rs.MoveNext
  Wend

End Sub

'mostra os dados da fonte de dados
Public Sub atualizaBaseDados(ByRef frm_importar As Form)

  Dim sSQL As String
  Dim rs As Recordset
  Dim listaTabelas(1 To 1000, 1 To 2) As String
  Dim k As Integer
  
  sSQL = " select distinct " & _
    "       datas.data_formatada, " & _
    "       BU, datas.datas " & _
    " from  " & _
    "       base_report_INFERIOR_t_bd  inner join datas on base_report_INFERIOR_t_bd.data = datas.datas order by datas.datas DESC"

  Set rs = CurrentDb.OpenRecordset(sSQL)

  While Not rs.EOF
    k = k + 1
    listaTabelas(k, 1) = rs.Fields("data_formatada").Value
    listaTabelas(k, 2) = rs.Fields("BU").Value
    rs.MoveNext
  Wend

  'preenche a msg de sucesso apos a importacao
  Dim accessUteis As accessUteis
  Set accessUteis = New accessUteis
  
  Call accessUteis.preencheListBox_2Columns(frm_importar.baseDados, listaTabelas, "Data:;BU:")
End Sub

'valida se o cabecalho contem os mesmos nomes parametrizados
Public Function validaTiposDeDados(DataRange As Variant, _
                                   datarangeNumerico As Variant, _
                                   cabecalho As Variant, _
                                   sheetName As String, _
                                   ultimaLinha As Double, _
                                   validacaoNumerica As Boolean) As Boolean
  Dim i As Double
  Dim valor As Variant
  
  'Verifica se existe algum erro de formula
  For Each valor In DataRange
    If IsError(valor) Then
      validaTiposDeDados = False 'se identificar algum erro de formula
      Exit Function
    End If
  Next valor
  '----------------------------------------------------

  If validacaoNumerica = False Then Exit Function

  '----------------------------------------------------
  'verifica se os dados sao realmente numericos
On Error GoTo trataErro
  For Each valor In datarangeNumerico
    
    'so valida valores preenchidos
    If valor <> "" Then
      
      'os dados so podem ser numericos
      If TypeName(CDbl(valor)) <> "Integer" And TypeName(CDbl(valor)) <> "Double" Then
        validaTiposDeDados = False 'se nao for numerico
        Exit Function
      End If
    End If
  Next valor
    
  validaTiposDeDados = True 'se estiver ok!

Exit Function
trataErro:
  validaTiposDeDados = False
End Function

'Monta o cabecalho de acordo com o relatorio
Public Function defineCabecalho(tipoFile As String) As Variant

  Dim cabecalho() As String
 
  If tipoFile = "Input Month Base" Or _
     tipoFile = "INPUT YTD Base Actual vs AOP" Or _
     tipoFile = "INPUT YTD Base Actual vs PY" Then
   
    ReDim cabecalho(1 To 24, 1 To 2)
    cabecalho(1, 1) = "SKU #"
    cabecalho(1, 2) = "texto"
    
    cabecalho(2, 1) = "SKU Description"
    cabecalho(2, 2) = "texto"
    
    cabecalho(3, 1) = "Category"
    cabecalho(3, 2) = "texto"
    
    cabecalho(4, 1) = "Regions/Channel"
    cabecalho(4, 2) = "texto"
    
    cabecalho(5, 1) = "BU"
    cabecalho(5, 2) = "texto"
    
    cabecalho(6, 1) = "Brand"
    cabecalho(6, 2) = "texto"
    
    cabecalho(7, 1) = ""
    
    cabecalho(8, 1) = "Tons"
    cabecalho(8, 2) = "numero"
    
    cabecalho(9, 1) = "LSV $"
    cabecalho(9, 2) = "numero"
    
    cabecalho(10, 1) = "NSV $"
    cabecalho(10, 2) = "numero"
    
    cabecalho(11, 1) = "VIC $"
    cabecalho(11, 2) = "numero"
    
    cabecalho(12, 1) = "VLC $"
    cabecalho(12, 2) = "numero"
    
    cabecalho(13, 1) = ""
    
    cabecalho(14, 1) = "Tons"
    cabecalho(14, 2) = "numero"
    
    cabecalho(15, 1) = "LSV $"
    cabecalho(15, 2) = "numero"
    
    cabecalho(16, 1) = "NSV $"
    cabecalho(16, 2) = "numero"
    
    cabecalho(17, 1) = "VIC $"
    cabecalho(17, 2) = "numero"
    
    cabecalho(18, 1) = "VLC $"
    cabecalho(18, 2) = "numero"
    
    cabecalho(19, 1) = ""
    
    cabecalho(20, 1) = "Tons"
    cabecalho(20, 2) = "numero"
    
    cabecalho(21, 1) = "LSV $"
    cabecalho(21, 2) = "numero"
    
    cabecalho(22, 1) = "NSV $"
    cabecalho(22, 2) = "numero"
    
    cabecalho(23, 1) = "VIC $"
    cabecalho(23, 2) = "numero"
    
    cabecalho(24, 1) = "VLC $"
    cabecalho(24, 2) = "numero"
  
  ElseIf tipoFile = "INPUT Scenario" Then
  
    ReDim cabecalho(1 To 4, 1 To 2)
    cabecalho(1, 1) = "BU"
    cabecalho(1, 2) = "texto"
  
    cabecalho(2, 1) = "Brand"
    cabecalho(2, 2) = "texto"
  
    cabecalho(3, 1) = "Regions/Channel"
    cabecalho(3, 2) = "texto"
  
    cabecalho(4, 1) = "Category Heinz"
    cabecalho(4, 2) = "texto"
  
  End If
  
  defineCabecalho = cabecalho 'return

End Function

'valida se o cabecalho contem os mesmos nomes parametrizados
Public Function validaCabecalhoNomes(datarangeCabecalho As Variant, _
                                     cabecalho As Variant) As Boolean
  Dim i As Integer
  For i = 1 To UBound(datarangeCabecalho, 2)
    If (datarangeCabecalho(1, i) <> cabecalho(1, i)) Then
      validaCabecalhoNomes = False 'se algum nome estiver alterado
      Exit Function
    End If
  Next i
  validaCabecalhoNomes = True 'se estiver ok!
End Function

'Cria o resultado da validacao
Public Function resultadoValidacao(ByRef lista As ListBox, _
                                         validacoes As Variant) As Variant

  Dim accessUteis As accessUteis
  Set accessUteis = New accessUteis
  
  Call accessUteis.preencheListBox_2Columns(lista, validacoes, "Validacao:;Status:")

End Function

'da a carga de dados na tabela 'base_report_INFERIOR_t'
Sub cargaEmBASE_report_INFERIOR()
  
  Dim rs As Recordset
  Set rs = CurrentDb.OpenRecordset("select * from base_report_INFERIOR")
  
  Dim coluna As Field
  Dim rs_tabelaLocal As Recordset
  Set rs_tabelaLocal = CurrentDb.OpenRecordset("select * from base_report_INFERIOR_t")
    
  'da a carga na tbale base_report_INFERIOR_t
  While Not rs.EOF
    rs_tabelaLocal.AddNew
    For Each coluna In rs_tabelaLocal.Fields
      coluna.Value = rs.Fields(coluna.name).Value
    Next coluna
    rs_tabelaLocal.Update
    rs.MoveNext 'proximo regitro da query base_report_INFERIOR
  Wend
  
End Sub

'limpa os dados do banco de dados
Public Sub limparBanco(ByRef frm_gerarReport As Form, _
                    Optional criterio As String)

  Dim tabelas(1 To 12) As String
  Dim i As Integer
  Dim criterioWhere As String
  
  'monta o criterio
  If criterio <> "" Then
    criterioWhere = " where " & criterio
  End If
  
  tabelas(1) = "base_report_INFERIOR_t_bd"
  tabelas(2) = "categoryMIX_report_final_AOP_adju_bd"
  tabelas(3) = "categoryMIX_report_final_AOP_bd"
  tabelas(4) = "categoryMIX_report_PY_adj_bd"
  tabelas(5) = "categoryMIX_report_PY_bd"
  tabelas(6) = "MIX_total_Categoria_Final_AOP_bd"
  tabelas(7) = "MIX_total_Categoria_Final_PY_II_final_bd"
  tabelas(8) = "regionMIX_report_bd"
  tabelas(9) = "regionMIX_report_PY_bd"
  tabelas(10) = "base_report_INFERIOR_t"
  tabelas(11) = "input_monthBase"
  tabelas(12) = "temp_input_monthBase"
  
  'faz a pergunta se deseja apagar a base de dados
  If MsgBox("Deseja excluir a base de dados?", vbYesNo) = vbNo Then
    Exit Sub
  End If
  
  DoCmd.SetWarnings False
  For i = 1 To UBound(tabelas)
    DoCmd.RunSQL "delete * from " & tabelas(i) & criterioWhere
  Next i
  DoCmd.SetWarnings True
  
  Call atualizaBaseDados(frm_gerarReport) 'mostra os dados da fonte de dados
End Sub
