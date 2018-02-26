Attribute VB_Name = "gerarReport"
Option Compare Database
Option Explicit

'cria um novo Excel com os dados da lista
Public Sub exportarLista(ByVal ssql As String, Optional reportIA As Boolean, Optional relatorio As String, Optional desabilitaTela As Boolean)
  Dim Load As Load
  Dim xls As Object
  Dim i As Long
  Dim campo As Object
  Dim rs As Recordset
  Set Load = New Load
  If ssql = "" Then Exit Sub
  Set rs = CurrentDb.OpenRecordset(ssql)
  Set xls = CreateObject("Excel.Application")
  '---------------------------
  xls.Application.Workbooks.Add 'adiciona um novo workbook
  '---------------------------
  xls.sheets(1).Range("A2").CopyFromRecordset rs 'cola os dados do recordset no workbook
  '---------------------------
  If desabilitaTela Then
    With xls.Application 'desabilita atualização de tela
      .Calculation = -4135 'xlCalculationManual
      .ScreenUpdating = False
      .EnableEvents = False
    End With
    xls.Visible = False
  Else
    xls.Visible = True
  End If
  '---------------------------
  'cola o cabecalho na linha 1
  For Each campo In rs.Fields
    i = i + 1
    xls.sheets(1).cells(1, i).Value = campo.Name
  Next campo
  '---------------------------
  If reportIA Then Call montaEstiloReport(xls) 'monta o estilo do report Excel'
  '---------------------------
  If relatorio = "ia_empresas" Then Call montaDicMarcasConcedidas(xls, relatorio) 'achuria as empresas
  If relatorio = "ia_empresas_contando" Then Call montaDicMarcasConcedidas(xls, relatorio) 'achuria as empresas
  '---------------------------
  With xls.Application 'habilita atualização de tela
    .Calculation = -4105 'xlCalculationAutomatic
    .ScreenUpdating = True
    .EnableEvents = True
  End With
  xls.Visible = True
  '---------------------------
  'insere as bordas na planilha
  With xls.ActiveSheet.UsedRange
    .Borders(7).ColorIndex = 0 'xlEdgeLeft
    .Borders(8).ColorIndex = 0 'xlEdgeTop
    .Borders(9).ColorIndex = 0 'xlEdgeBottom
    .Borders(10).ColorIndex = 0 'xlEdgeRight
    .Borders(11).ColorIndex = 0 'xlInsideVertical
    .Borders(12).ColorIndex = 0 'xlInsideHorizontal
  End With
End Sub


'preenche o relatorio com os tipos
Public Sub preencherListaReport(ByRef lista As ListBox)
  
  Dim relatorios(1 To 6) As String

  relatorios(1) = "INPUT Scenario"
  relatorios(2) = "BRAIN vs AOP"
  relatorios(3) = "BRAIN vs PY"
  relatorios(4) = "Outputs LATAM"
  relatorios(5) = "Outputs MPR"
  relatorios(6) = "Scorecards"

  Dim accessUteis As accessUteis
  Set accessUteis = New accessUteis
  
  '---------------------------------------------
  'Cria o relatorio com os tipos
  Call accessUteis.preencheListBox_1Columns(lista, relatorios, "Tipos")

End Sub

'Seleciona o Excel que sera importado
Public Function SelecionaPasta(ByRef pasta_endereco As TextBox) As String
  
  Dim FileDialogs As FileDialogs
  Set FileDialogs = New FileDialogs
  
  pasta_endereco.Value = FileDialogs.SelectFolderFile()
  
End Function

'gera os relatorios do app
Public Sub gerarRelatorios(ByRef frm_gerarReport As Form)

  Dim rsRelatorioAgrupado As Recordset

  Dim inicioPros As Double
  inicioPros = Timer
  
  Dim CriandoObjetos As CriandoObjetos
  Set CriandoObjetos = New CriandoObjetos
  
  Dim InstanciaExcel As InstanciaExcel
  Set InstanciaExcel = New InstanciaExcel
  
  Dim accessUteis As accessUteis
  Set accessUteis = New accessUteis
  
  Dim pasta As String
  pasta = frm_gerarReport.pasta_endereco.Value
  
  Dim COnverter As COnverter
  Set COnverter = New COnverter
  
  Dim dataInicio As String
  dataInicio = COnverter.isVazio(frm_gerarReport.dataInicial.Value)
  Dim dataFinal As String
  dataFinal = COnverter.isVazio(frm_gerarReport.dataFinal.Value)
  Dim relatorios As Variant
  relatorios = accessUteis.getListValues(frm_gerarReport.lista_reports, 0)
  Dim relatorioMensal As Boolean
  relatorioMensal = frm_gerarReport.relatorioMensal.Value
  Dim relatorioAgrupado As Boolean
  relatorioAgrupado = frm_gerarReport.relatorioAgrupado.Value
  Dim chk_BU As Boolean
  chk_BU = frm_gerarReport.chk_BU.Value
  Dim cbo_BU As String
  cbo_BU = frm_gerarReport.cbo_BU.Value
  
  Dim Load As Load
  Set Load = New Load
    
  '---------------------------------------------------------------------
  'Valida se o nome do relatório é válido
  If pasta = "" Then
    MsgBox "Selecione um endereço válido para salvar o arquivo", vbCritical
    Exit Sub
  End If
  
  'Valida se existe o arquivo na pasta
  If CriandoObjetos.ArquivoExiste(pasta) = True Then
    MsgBox "Já existe um arquivo no local: '" & pasta & "'", vbCritical
    Exit Sub
  End If
  
  '---------------------------------------------------------------------
  'verifica se o tipo de relatório, data e o arquivo foram selecionados
  If ((pasta = "") Or _
     (Len(Join(relatorios)) = 0) Or _
     (dataFinal = "") Or _
     (dataInicio = "")) Then
    MsgBox "É necessário selecionar o 'tipo' de relatório, a 'data' e a 'pasta' que será gerado.", vbCritical, "Selecione os campos"
    Exit Sub
  End If
  
  'se o check de BU estiver ativo, e' necessario seleciona a bu
  If chk_BU And cbo_BU = "" Then
    MsgBox "É necessário selecionar a 'BU'.", vbCritical, "Selecione a 'BU'"
    Exit Sub
  End If
  
'  'Verifica se o relatorio mensal ou agrupado foi selecionado
'  If dataInicio = "" And (itemSelecionado(frm_gerarReport.lista_reports, "Outputs LATAM") = 1 Or itemSelecionado(frm_gerarReport.lista_reports, "Outputs MPR") = 1) Then
'    MsgBox "O relatorio de 'Outputs' precisa de 'data-inicio' e 'data-fim'.", vbCritical, "Selecione os campos"
'    Exit Sub
'  End If
  
  
'  'Verifica se o relatorio mensal ou agrupado foi selecionado
'  If (relatorioMensal = False And relatorioAgrupado = False) And maisItensSelecionadosAlemDeOutputs(frm_gerarReport.lista_reports) = 1 Then
'    MsgBox "É necessário selecionar o tipo de relatório: 'Mensal' ou 'Agrupado'", vbCritical, "Selecione os campos"
'    Exit Sub
'  End If
  
'  'Verifica se o relatorio mensal ou agrupado foi selecionado
'  If (relatorioMensal = True Or relatorioAgrupado = True) And maisItensSelecionadosAlemDeOutputs(frm_gerarReport.lista_reports) = 0 Then
'    MsgBox "O relatório a 'mensal' ou 'agrupado' foram selecionados. Selecione um tipo de relatorio valido.", vbCritical, "Selecione os campos"
'    Exit Sub
'  End If
  
  '---------------------------------------------------------------------
'  'Se for selecionado o relatório agrupado, tem que ter data de inicio
'  If (relatorioAgrupado And dataInicio = "") Then
'    MsgBox "É necessário selecionar a 'Data Inicial' do relatório", vbCritical, "Selecione os campos"
'    Exit Sub
'  End If
    
  '---------------------------------------------------------------------
  
  '++++++++++++++++++++++++++++++++++++++++++
'  If dataInicio = "" Then
'    dataInicio = dataFinal
'  End If
  '++++++++++++++++++++++++++++++++++++++++++
  
  'a data Inicial nao pode ser maior que a data final
  If (CDate(dataInicio) > CDate(dataFinal)) Then
    MsgBox "A data inicial tem que ser menor ou igual que a data inicial", vbCritical, "Selecione os campos"
    Exit Sub
  End If
    
  '---------------------------------------------------------------------
  'Exclui a base de dados anterior e da a carga na nova
  DoCmd.SetWarnings False
  Call InstanciaExcel.NewFile 'cria um novo arquivo de Excel
  'Call InstanciaExcel.CriaQTDSheet(UBound(relatorios) + IIf(relatorioAgrupado And relatorioMensal, 2, 0)) '<---- adiciona os 2 relatorio agrupados
  
  '---------------------------------------------------------------------
  'faz um iteração pelo array montando as sheets e os relatórios
  Dim rs As Recordset
  
  'valida se ha dados
  Set rs = CurrentDb.OpenRecordset("select * from base_report_INFERIOR_t_bd where data = #" & Format(dataFinal, "mm/dd/yyyy") & "# ")
  If rs.RecordCount = 0 Then
    MsgBox "Não existem dados na base de dados do relatório para a data: '" & dataFinal & "'", vbCritical
    Exit Sub
  End If

  Call insereNaFiltroAgrupado(dataInicio, dataInicio) 'para o mensal de bu_consolidated

  'se for agrupado verifica se ha dados em todas as datas
  If relatorioAgrupado Then
    Call insereNaFiltroAgrupado(dataInicio, dataFinal)
    Set rsRelatorioAgrupado = CurrentDb.OpenRecordset("select distinct data, filtraAgrupado.dataFiltro from base_report_INFERIOR_t_bd right join filtraAgrupado on base_report_INFERIOR_t_bd.data =  filtraAgrupado.dataFiltro where data is null")
  
    If rsRelatorioAgrupado.RecordCount > 0 Then
      MsgBox "Não existem dados na base de dados do relatório para a data: '" & rsRelatorioAgrupado.Fields("dataFiltro").Value & "'", vbCritical
      Exit Sub
    End If
  End If


  '---------------------------------------------------------------------
  '********************************************
  'se for a 'BRAIN' ou INPUT
  If itemSelecionado(frm_gerarReport.lista_reports, "BRAIN") = 1 Or itemSelecionado(frm_gerarReport.lista_reports, "INPUT") = 1 Then
    Call gerarReports(relatorios, InstanciaExcel, dataInicio, dataFinal, False, frm_gerarReport, "Actual", chk_BU, cbo_BU)
    Call gerarReports(relatorios, InstanciaExcel, dataInicio, dataFinal, True, frm_gerarReport, "YTD", chk_BU, cbo_BU)
    Call formataReports(InstanciaExcel.xls, "Brain", InstanciaExcel, dataFinal) 'formata reports gerados
  End If
  
  '********************************************
  'se for o LATAM
  If itemSelecionado(frm_gerarReport.lista_reports, "Outputs LATAM") = 1 Then
      
    'se for agrupado verifica se ha dados em todas as datas
    Call insereNaFiltroAgrupado(dataInicio, dataFinal)
    Set rsRelatorioAgrupado = CurrentDb.OpenRecordset("select distinct data, filtraAgrupado.dataFiltro from base_report_INFERIOR_t_bd right join filtraAgrupado on base_report_INFERIOR_t_bd.data =  filtraAgrupado.dataFiltro where data is null")
    
    If rsRelatorioAgrupado.RecordCount > 0 Then
      MsgBox "Não existem dados na base de dados do relatório para a data: '" & rsRelatorioAgrupado.Fields("dataFiltro").Value & "'", vbCritical
      Exit Sub
    End If
    
    Call gerarReports(relatorios, InstanciaExcel, dataInicio, dataFinal, False, frm_gerarReport, "Outputs LATAM", chk_BU, cbo_BU)
    Call formataReports(InstanciaExcel.xls, "Output", InstanciaExcel, dataFinal) 'formata reports gerados
  End If
    
  '********************************************
  'se for o MPR
  If itemSelecionado(frm_gerarReport.lista_reports, "Outputs MPR") = 1 Then
      
    'se for agrupado verifica se ha dados em todas as datas
    Call insereNaFiltroAgrupado(dataInicio, dataFinal)
    Set rsRelatorioAgrupado = CurrentDb.OpenRecordset("select distinct data, filtraAgrupado.dataFiltro from base_report_INFERIOR_t_bd right join filtraAgrupado on base_report_INFERIOR_t_bd.data =  filtraAgrupado.dataFiltro where data is null")
    
    If rsRelatorioAgrupado.RecordCount > 0 Then
      MsgBox "Não existem dados na base de dados do relatório para a data: '" & rsRelatorioAgrupado.Fields("dataFiltro").Value & "'", vbCritical
      Exit Sub
    End If
    
    Call gerarReports(relatorios, InstanciaExcel, dataInicio, dataFinal, False, frm_gerarReport, "Outputs MPR", chk_BU, cbo_BU)
    Call formataReports(InstanciaExcel.xls, "MPR", InstanciaExcel, dataFinal) 'formata reports gerados
  End If


  '********************************************
  'se for o scorecards
  If itemSelecionado(frm_gerarReport.lista_reports, "Scorecards") = 1 Then
      
    'se for agrupado verifica se ha dados em todas as datas
    Call insereNaFiltroAgrupado(dataInicio, dataFinal)
    Set rsRelatorioAgrupado = CurrentDb.OpenRecordset("select distinct data, filtraAgrupado.dataFiltro from base_report_INFERIOR_t_bd right join filtraAgrupado on base_report_INFERIOR_t_bd.data =  filtraAgrupado.dataFiltro where data is null")
    
    If rsRelatorioAgrupado.RecordCount > 0 Then
      MsgBox "Não existem dados na base de dados do relatório para a data: '" & rsRelatorioAgrupado.Fields("dataFiltro").Value & "'", vbCritical
      Exit Sub
    End If
    
    Call gerarReports(relatorios, InstanciaExcel, dataInicio, dataFinal, False, frm_gerarReport, "Scorecards", chk_BU, cbo_BU)
    Call formataReports(InstanciaExcel.xls, "Scorecards", InstanciaExcel, dataFinal) 'formata reports gerados
  End If

  '---------------------------------------------------------------------
  
  Call InstanciaExcel.SaveAs(pasta)
  Set InstanciaExcel = Nothing
  Set Load = Nothing
  
  Debug.Print inicioPros - Timer
  MsgBox "Arquivo criado em '" & pasta & "'", vbInformation

End Sub


'Gerar relatórios
Public Sub gerarReports(ByVal relatorios As Variant, _
                        ByRef InstanciaExcel As InstanciaExcel, _
                              dataInicio As String, _
                              dataFinal As String, _
                              isAgrupado As Boolean, _
                      ByRef frm_gerarReport As Form, _
                              outputsLatam As String, _
                              chk_BU As Boolean, _
                              cbo_BU As String)

  Dim sqlReport As String
  Dim k As Integer
  Dim nomeReport As String
  Dim rs As Recordset
  Dim brain_actual_vs_AOP_reports(1 To 8, 1 To 6) As String
  Dim ii As Integer
  Dim reportNameAgrupado As String
  Dim queryStringAgrupada As DAO.QueryDef
  Dim reportOutputs(1 To 4, 1 To 5) As String
  Dim j As Integer
  Dim sSQLdistinctCategoria As String
  Dim rsDistinct As Recordset
  Dim rsTotal As Recordset
  Dim rsDetalhe As Recordset
  Dim rsBU As Recordset
  Dim rsRegionChannel As Recordset
  Dim rsRegionsChannelReal As Recordset
  Dim y As Integer
  Dim reportScorecards(1 To 12, 1 To 7) As String
  Dim BU_filtro As String

  'pega a dataInicial e dataFinal e insere na filtro agrupado
  Call insereNaFiltroAgrupado(dataInicio, dataFinal)
  
  'se o report de BU for selecionado
  If chk_BU Then
    BU_filtro = " and BU='" & cbo_BU & "'"
    Call insereFiltroBU(cbo_BU) 'insere filtro de BU
  Else
    Call nenhumNaFiltroBU 'insere tudo na BU
  End If

  '-----------------------------------------------------------
  'se o report de agrupado for selecionado
  If isAgrupado Then
    reportNameAgrupado = "_agrupado"
  End If
  
  '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  For k = 1 To UBound(relatorios)
    
    'nao adiciona novas 'INPUT Scenario'
    If InstanciaExcel.ValidaSeSheetExiste("INPUT Scenario") And relatorios(k) = "INPUT Scenario" Then GoTo nextIteration
    
    'so adiciona sheets que forem novas e nao estiverem preenchidas
    If InstanciaExcel.xls.Sheets(1).UsedRange.Rows.Count > 1 And (relatorios(k) <> "Outputs LATAM" And outputsLatam <> "Outputs LATAM") And (relatorios(k) <> "Outputs MPR" And outputsLatam <> "Outputs MPR") And (relatorios(k) <> "Scorecards" And outputsLatam <> "Scorecards") Then
      InstanciaExcel.xls.Sheets.Add After:=InstanciaExcel.xls.Sheets(InstanciaExcel.xls.Sheets.Count)  'adiciona um sheet a cada iteracao
    End If
    
    '---------------------------------------------------------------------
    If relatorios(k) = "INPUT Scenario" And (outputsLatam <> "Outputs LATAM" And outputsLatam <> "Outputs MPR" And outputsLatam <> "Scorecards") Then
      nomeReport = "INPUT Scenario"
      
      InstanciaExcel.xls.Sheets(InstanciaExcel.xls.Sheets.Count).name = nomeReport
                  sqlReport = "select distinct  "
      sqlReport = sqlReport & " data, "
      sqlReport = sqlReport & " '' as BU_, "
      sqlReport = sqlReport & " '' as Brand_, "
      sqlReport = sqlReport & " '' as Regions_Channel_, "
      sqlReport = sqlReport & " '' as Category_ "
      sqlReport = sqlReport & "from "
      sqlReport = sqlReport & " base_report_INFERIOR_t_bd "
      sqlReport = sqlReport & "where  "
      sqlReport = sqlReport & " data between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro
      sqlReport = sqlReport & "union all "
      sqlReport = sqlReport & "select distinct  "
      sqlReport = sqlReport & " '', "
      sqlReport = sqlReport & " BU, "
      sqlReport = sqlReport & " '', "
      sqlReport = sqlReport & " '', "
      sqlReport = sqlReport & " '' "
      sqlReport = sqlReport & "from  "
      sqlReport = sqlReport & " base_report_INFERIOR_t_bd "
      sqlReport = sqlReport & "where  "
      sqlReport = sqlReport & " data between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro
      sqlReport = sqlReport & "union all "
      sqlReport = sqlReport & "select distinct  "
      sqlReport = sqlReport & " '', "
      sqlReport = sqlReport & " '',  "
      sqlReport = sqlReport & " Brand, "
      sqlReport = sqlReport & " '', "
      sqlReport = sqlReport & " '' "
      sqlReport = sqlReport & "from  "
      sqlReport = sqlReport & " base_report_INFERIOR_t_bd "
      sqlReport = sqlReport & "where  "
      sqlReport = sqlReport & " data between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# and Brand <> '0' " & BU_filtro
      sqlReport = sqlReport & "union all "
      sqlReport = sqlReport & "select distinct  "
      sqlReport = sqlReport & " '', "
      sqlReport = sqlReport & " '',  "
      sqlReport = sqlReport & " '',  "
      sqlReport = sqlReport & " Regions_Channel, "
      sqlReport = sqlReport & " '' "
      sqlReport = sqlReport & "from  "
      sqlReport = sqlReport & " base_report_INFERIOR_t_bd "
      sqlReport = sqlReport & "where  "
      sqlReport = sqlReport & " data between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# and Regions_Channel <> '0' " & BU_filtro
      sqlReport = sqlReport & "union all "
      sqlReport = sqlReport & "select distinct  "
      sqlReport = sqlReport & " '', "
      sqlReport = sqlReport & " '',  "
      sqlReport = sqlReport & " '', "
      sqlReport = sqlReport & " '',  "
      sqlReport = sqlReport & " Category "
      sqlReport = sqlReport & "from  "
      sqlReport = sqlReport & " base_report_INFERIOR_t_bd "
      sqlReport = sqlReport & "where  "
      sqlReport = sqlReport & " data between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# and Category <> '0' " & BU_filtro
  
      'cola os dados no report '* <----------------------
      Set rs = CurrentDb.OpenRecordset(sqlReport)
      If rs.RecordCount = 0 Then
        MsgBox "Não existem dados na base de dados do relatório '" & nomeReport & "': " & dataFinal, vbCritical
        Exit Sub
      End If
      
      Call InstanciaExcel.insereTitulo(nomeReport, "INPUT Scenario", False)
      Call InstanciaExcel.ColaDadosRS(rs, nomeReport, "B", 2)
      Call InstanciaExcel.converteRangeTabela(nomeReport, nomeReport, True, "B", 2, False) 'converte o range colado em tabela
      'Call InstanciaExcel.ZoomOitenta(nomeReport)   'Muda a config do Excel para o tamanho zoom: 80
      'Call InstanciaExcel.formataPrimeiraColunaData(nomeReport) 'formata a primeira coluna da sheet como data
      
    '---------------------------------------------------------------------
    ElseIf relatorios(k) = "BRAIN vs AOP" And (outputsLatam <> "Outputs LATAM" And outputsLatam <> "Outputs MPR" And outputsLatam <> "Scorecards") Then
      nomeReport = "BRAIN Actual vs AOP"
      If reportNameAgrupado <> "" Then
        nomeReport = "BRAIN YTD vs AOP"
      End If
      
      InstanciaExcel.xls.Sheets(InstanciaExcel.xls.Sheets.Count).name = nomeReport
      
      'montagem do relatório em 3 partes
      '1 - MIX_total_Categoria_Final
      '2 - categoryMIX_report_final
      '3 - base_report_INFERIOR_t

      'BU consolidado
      brain_actual_vs_AOP_reports(1, 1) = "bu_consolidate_AOP_adj_agrupado_union_a"
      brain_actual_vs_AOP_reports(1, 2) = "BU Consolidated"
      If chk_BU Then brain_actual_vs_AOP_reports(1, 3) = " where BU='" & cbo_BU & "'"
      If chk_BU Then brain_actual_vs_AOP_reports(1, 4) = " where BU='" & cbo_BU & "'"
      brain_actual_vs_AOP_reports(1, 5) = "B" 'coluna inicial
      brain_actual_vs_AOP_reports(1, 6) = 2 'coluna inicial

      brain_actual_vs_AOP_reports(2, 1) = "bu_consolidate_AOP_adj_union_a" & reportNameAgrupado
      brain_actual_vs_AOP_reports(2, 2) = "Month Consolidated"
      brain_actual_vs_AOP_reports(2, 3) = " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro
      If chk_BU Then brain_actual_vs_AOP_reports(2, 4) = " where BU='" & cbo_BU & "'"
      brain_actual_vs_AOP_reports(2, 5) = "B" 'coluna inicial
      brain_actual_vs_AOP_reports(2, 6) = 2 'coluna inicial
        
'      brain_actual_vs_AOP_reports(2, 1) = "month_consolidate_AOP" & reportNameAgrupado
'      brain_actual_vs_AOP_reports(2, 2) = "Month Consolidated"
'      brain_actual_vs_AOP_reports(2, 3) = " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & " order by Month asc, BU asc, Category asc" 'nao agrupado
'      brain_actual_vs_AOP_reports(2, 4) = " order by BU asc"  'agrupado do mensal
'      brain_actual_vs_AOP_reports(2, 5) = "B" 'coluna inicial
'      brain_actual_vs_AOP_reports(2, 6) = 2 'coluna inicial
       
      'MIX_total_Categoria_Final_AOP_bd_a_adju -->> com (adju)
      brain_actual_vs_AOP_reports(3, 1) = "MIX_total_Categoria_Final_AOP_bd_a" & reportNameAgrupado
      brain_actual_vs_AOP_reports(3, 2) = "Category Consolidated"
      brain_actual_vs_AOP_reports(3, 3) = " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# and Category<>'0' " & BU_filtro & " order by Month asc, BU asc, Category asc"   'nao agrupado
      brain_actual_vs_AOP_reports(3, 4) = " having MIX_total_Categoria_Final_AOP_bd.Category<>'0' order by MIX_total_Categoria_Final_AOP_bd.BU asc, MIX_total_Categoria_Final_AOP_bd.Category asc"   'agrupado do mensal
      brain_actual_vs_AOP_reports(3, 5) = "B" 'coluna inicial
      brain_actual_vs_AOP_reports(3, 6) = 2 'coluna inicial
      
      brain_actual_vs_AOP_reports(4, 1) = "region_consolidate_AOP" & reportNameAgrupado
      brain_actual_vs_AOP_reports(4, 2) = "Region Consolidated"
      brain_actual_vs_AOP_reports(4, 3) = " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# and [Region/Channel] <> '0'" & BU_filtro & " and Category<>'0' order by Month asc, BU asc, [Region/Channel] asc"  'nao agrupa
      brain_actual_vs_AOP_reports(4, 4) = " having categoryMIX_report_final_AOP_bd.Regions_Channel <> '0' " & IIf(chk_BU, " and categoryMIX_report_final_AOP_bd.BU='" & cbo_BU & "'", "") & " order by categoryMIX_report_final_AOP_bd.BU asc"   'agrupado do mensal
      brain_actual_vs_AOP_reports(4, 5) = "B" 'coluna inicial
      brain_actual_vs_AOP_reports(4, 6) = 2 'coluna inicial
      
      brain_actual_vs_AOP_reports(5, 1) = "categoryMIX_report_final_AOP_bd_a" & reportNameAgrupado
      brain_actual_vs_AOP_reports(5, 2) = "Region/Category Consolidated"
      brain_actual_vs_AOP_reports(5, 3) = " where Category<>'0' and Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro & " and ([Region/Channel] <> '0') order by Month asc, BU asc, [Region/Channel] asc, [Region/Channel] asc, Category asc"
      brain_actual_vs_AOP_reports(5, 4) = " having categoryMIX_report_final_AOP_bd.[Category]<>'0' and  categoryMIX_report_final_AOP_bd.[Regions_Channel] <> '0' " & IIf(chk_BU, " and categoryMIX_report_final_AOP_bd.BU='" & cbo_BU & "'", "") & " order by categoryMIX_report_final_AOP_bd.BU asc, categoryMIX_report_final_AOP_bd.[Regions_Channel] asc, categoryMIX_report_final_AOP_bd.Category asc"
      brain_actual_vs_AOP_reports(5, 5) = "A" 'coluna inicial
      brain_actual_vs_AOP_reports(5, 6) = 1 'coluna inicial
      
      brain_actual_vs_AOP_reports(6, 1) = "categoryMIX_report_final_AOP_adju_bd_a_BU" & reportNameAgrupado
      brain_actual_vs_AOP_reports(6, 2) = "Adju/BU Consolidated"
      brain_actual_vs_AOP_reports(6, 3) = " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro & " and ([Region/Channel] <> '0') order by Month asc, BU asc, [Region/Channel] asc, [Region/Channel] asc, Category asc"
      brain_actual_vs_AOP_reports(6, 4) = IIf(chk_BU, " having categoryMIX_report_final_AOP_adju_bd.BU='" & cbo_BU & "'", "") & "  order by categoryMIX_report_final_AOP_adju_bd.BU asc "
      brain_actual_vs_AOP_reports(6, 5) = "B" 'coluna inicial
      brain_actual_vs_AOP_reports(6, 6) = 2 'coluna inicial
    
      brain_actual_vs_AOP_reports(7, 1) = "categoryMIX_report_final_AOP_adju_bd_a" & reportNameAgrupado
      brain_actual_vs_AOP_reports(7, 2) = "Adju/Region Consolidated"
      brain_actual_vs_AOP_reports(7, 3) = " where Category<>'0' and Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro & " and ([Region/Channel] <> '0') order by Month asc, BU asc, [Region/Channel] asc, [Region/Channel] asc, Category asc"
      brain_actual_vs_AOP_reports(7, 4) = " having categoryMIX_report_final_AOP_adju_bd.Category<>'0' " & IIf(chk_BU, " and categoryMIX_report_final_AOP_adju_bd.BU='" & cbo_BU & "'", "") & "  order by categoryMIX_report_final_AOP_adju_bd.BU asc, categoryMIX_report_final_AOP_adju_bd.Category asc"
      brain_actual_vs_AOP_reports(7, 5) = "B" 'coluna inicial
      brain_actual_vs_AOP_reports(7, 6) = 2 'coluna inicial
      
      brain_actual_vs_AOP_reports(8, 1) = "base_report_INFERIOR_t_AOP_a" & reportNameAgrupado
      brain_actual_vs_AOP_reports(8, 2) = "Input Base"
      brain_actual_vs_AOP_reports(8, 3) = " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro & "  and ([Region/Channel] <> '0' and SKU <> '0') order by BU, Month asc, SKU asc"
      brain_actual_vs_AOP_reports(8, 4) = " having Regions_Channel <> '0' " & IIf(chk_BU, " and BU='" & cbo_BU & "'", "") & " order by BU, SKU asc, Category asc "
      brain_actual_vs_AOP_reports(8, 5) = "A" 'coluna inicial
      brain_actual_vs_AOP_reports(8, 6) = 1 'coluna inicial
      
      'cola cada parte das queries na sheet
      For ii = 1 To UBound(brain_actual_vs_AOP_reports, 1)
  
        'se for agrupado
        If isAgrupado Then
          Set queryStringAgrupada = CurrentDb.QueryDefs(brain_actual_vs_AOP_reports(ii, 1))
          sqlReport = queryStringAgrupada.SQL & brain_actual_vs_AOP_reports(ii, 4)
          sqlReport = Replace(sqlReport, ";", "")
        
        'se nao for agrupado (mensal)
        Else
          sqlReport = "select * from " & brain_actual_vs_AOP_reports(ii, 1) & brain_actual_vs_AOP_reports(ii, 3)
        End If
        
        'cola os dados no report '* <----------------------
        Set rs = CurrentDb.OpenRecordset(sqlReport)
        If rs.RecordCount = 0 Then
          MsgBox "Não existem dados na base de dados do relatório '" & nomeReport & "': " & dataFinal, vbCritical
          Exit Sub
        End If
        
        Call InstanciaExcel.insereTitulo(nomeReport, brain_actual_vs_AOP_reports(ii, 2), False)
        Call InstanciaExcel.ColaDadosRS(rs, nomeReport, brain_actual_vs_AOP_reports(ii, 5), brain_actual_vs_AOP_reports(ii, 6))  'cola os dados do recordset na planilha | formata o estilo da header aqui
        Call InstanciaExcel.converteRangeTabela(nomeReport, nomeReport & ii, True, brain_actual_vs_AOP_reports(ii, 5), brain_actual_vs_AOP_reports(ii, 6), False) 'converte o range colado em tabela
        
        'If isAgrupado = False Then Call InstanciaExcel.formataPrimeiraColunaData(nomeReport) 'formata a primeira coluna da sheet como data
      Next ii
              
    '---------------------------------------------------------------------
    ElseIf relatorios(k) = "BRAIN vs PY" And (outputsLatam <> "Outputs LATAM" And outputsLatam <> "Outputs MPR" And outputsLatam <> "Scorecards") Then
      nomeReport = "BRAIN Actual vs PY" & reportNameAgrupado
      If reportNameAgrupado <> "" Then
        nomeReport = "BRAIN YTD vs PY"
      End If
      InstanciaExcel.xls.Sheets(InstanciaExcel.xls.Sheets.Count).name = nomeReport
            
      'BU consolidado
      brain_actual_vs_AOP_reports(1, 1) = "bu_consolidate_PY_adj_agrupado_union_a"
      brain_actual_vs_AOP_reports(1, 2) = "BU Consolidated"
      If chk_BU Then brain_actual_vs_AOP_reports(1, 3) = " where BU='" & cbo_BU & "'"
      If chk_BU Then brain_actual_vs_AOP_reports(1, 4) = " where BU='" & cbo_BU & "'"
      brain_actual_vs_AOP_reports(1, 5) = "B" 'coluna inicial
      brain_actual_vs_AOP_reports(1, 6) = 2 'coluna inicial
      
      brain_actual_vs_AOP_reports(2, 1) = "bu_consolidate_PY_adj_union_a" & reportNameAgrupado
      brain_actual_vs_AOP_reports(2, 2) = "Month Consolidated"
      brain_actual_vs_AOP_reports(2, 3) = " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro
      If chk_BU Then brain_actual_vs_AOP_reports(2, 4) = " where BU='" & cbo_BU & "'"
      brain_actual_vs_AOP_reports(2, 5) = "B" 'coluna inicial
      brain_actual_vs_AOP_reports(2, 6) = 2 'coluna inicial
           
'      brain_actual_vs_AOP_reports(2, 1) = "month_consolidate_PY" & reportNameAgrupado
'      brain_actual_vs_AOP_reports(2, 2) = "Month Consolidated"
'      brain_actual_vs_AOP_reports(2, 3) = " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & " order by Month asc, BU asc, Category asc" 'nao agrupado
'      brain_actual_vs_AOP_reports(2, 4) = " order by BU asc"  'agrupado do mensal
'      brain_actual_vs_AOP_reports(2, 5) = "B" 'coluna inicial
'      brain_actual_vs_AOP_reports(2, 6) = 2 'coluna inicial

      brain_actual_vs_AOP_reports(3, 1) = "MIX_total_Categoria_Final_PY_II_final_bd_a" & reportNameAgrupado
      brain_actual_vs_AOP_reports(3, 2) = "Category Consolidated"
      brain_actual_vs_AOP_reports(3, 3) = " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro & " and Category<>'0' order by Month asc, BU asc, Category asc"   'nao agrupado
      brain_actual_vs_AOP_reports(3, 4) = " having  MIX_total_Categoria_Final_PY_II_final_bd.Category<>'0' " & IIf(chk_BU, " and MIX_total_Categoria_Final_PY_II_final_bd.BU='" & cbo_BU & "'", "") & " order by MIX_total_Categoria_Final_PY_II_final_bd.BU asc, MIX_total_Categoria_Final_PY_II_final_bd.Category asc" 'agrupado do mensal
      brain_actual_vs_AOP_reports(3, 5) = "B" 'coluna inicial
      brain_actual_vs_AOP_reports(3, 6) = 2 'coluna inicial
      
      brain_actual_vs_AOP_reports(4, 1) = "region_consolidate_PY" & reportNameAgrupado
      brain_actual_vs_AOP_reports(4, 2) = "Region Consolidated"
      brain_actual_vs_AOP_reports(4, 3) = " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro & " and [Region/Channel] <> '0'" & " order by Month asc, BU asc, [Region/Channel] asc"  'nao agrupa
      brain_actual_vs_AOP_reports(4, 4) = " having categoryMIX_report_PY_bd.Regions_Channel <> '0' " & IIf(chk_BU, " and categoryMIX_report_PY_bd.BU='" & cbo_BU & "'", "") & " order by categoryMIX_report_PY_bd.BU asc"      'agrupado do mensal
      brain_actual_vs_AOP_reports(4, 5) = "B" 'coluna inicial
      brain_actual_vs_AOP_reports(4, 6) = 2 'coluna inicial
                 
      brain_actual_vs_AOP_reports(5, 1) = "categoryMIX_report_PY_bd_a" & reportNameAgrupado
      brain_actual_vs_AOP_reports(5, 2) = "Region/Category Consolidated"
      brain_actual_vs_AOP_reports(5, 3) = " where Category<>'0' and Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro & " and ([Region/Channel] <> '0') order by Month asc, BU asc, [Region/Channel] asc, [Region/Channel] asc, Category asc"
      brain_actual_vs_AOP_reports(5, 4) = " having Category<>'0' and [Regions_Channel] <> '0' " & IIf(chk_BU, " and BU='" & cbo_BU & "'", "") & " order by BU asc, [Regions_Channel] asc, Category asc"
      brain_actual_vs_AOP_reports(5, 5) = "A" 'coluna inicial
      brain_actual_vs_AOP_reports(5, 6) = 1 'coluna inicial
      
      brain_actual_vs_AOP_reports(6, 1) = "categoryMIX_report_PY_adj_bd_a_BU" & reportNameAgrupado
      brain_actual_vs_AOP_reports(6, 2) = "Adju/BU Consolidated"
      brain_actual_vs_AOP_reports(6, 3) = " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro & " and ([Region/Channel] <> '0') order by Month asc, BU asc, [Region/Channel] asc, [Region/Channel] asc, Category asc"
      brain_actual_vs_AOP_reports(6, 4) = IIf(chk_BU, " having categoryMIX_report_PY_adj_bd.BU='" & cbo_BU & "'", "") & " order by BU asc "
      brain_actual_vs_AOP_reports(6, 5) = "B" 'coluna inicial
      brain_actual_vs_AOP_reports(6, 6) = 2 'coluna inicial
      
      brain_actual_vs_AOP_reports(7, 1) = "categoryMIX_report_PY_adj_bd_a" & reportNameAgrupado
      brain_actual_vs_AOP_reports(7, 2) = "Adju/Region Consolidated"
      brain_actual_vs_AOP_reports(7, 3) = " where Category<>'0' and Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro & " and ([Region/Channel] <> '0') order by Month asc, BU asc, [Region/Channel] asc, [Region/Channel] asc, Category asc"
      brain_actual_vs_AOP_reports(7, 4) = " having categoryMIX_report_PY_adj_bd.Category<>'0' " & IIf(chk_BU, " and categoryMIX_report_PY_adj_bd.BU='" & cbo_BU & "'", "") & " order by BU asc, Category asc"
      brain_actual_vs_AOP_reports(7, 5) = "B" 'coluna inicial
      brain_actual_vs_AOP_reports(7, 6) = 2 'coluna inicial
      
      brain_actual_vs_AOP_reports(8, 1) = "base_report_INFERIOR_t_PY_a" & reportNameAgrupado
      brain_actual_vs_AOP_reports(8, 2) = "Input Base"
      brain_actual_vs_AOP_reports(8, 3) = " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# " & BU_filtro & "  and ([Region/Channel] <> '0' and SKU <> '0') order by BU, Month asc, SKU asc"
      brain_actual_vs_AOP_reports(8, 4) = " having  Regions_Channel <> '0' " & IIf(chk_BU, " and BU='" & cbo_BU & "'", "") & " order by BU, SKU asc, Category asc "
      brain_actual_vs_AOP_reports(8, 5) = "A" 'coluna inicial
      brain_actual_vs_AOP_reports(8, 6) = 1 'coluna inicial
      
      'cola cada parte das queries na sheet
      For ii = 1 To UBound(brain_actual_vs_AOP_reports, 1)
        
        'se for agrupado
        If isAgrupado Then
          Set queryStringAgrupada = CurrentDb.QueryDefs(brain_actual_vs_AOP_reports(ii, 1))
          sqlReport = queryStringAgrupada.SQL & brain_actual_vs_AOP_reports(ii, 4)
          sqlReport = Replace(sqlReport, ";", "")
        
        'se nao for agrupado (mensal)
        Else
          sqlReport = "select * from " & brain_actual_vs_AOP_reports(ii, 1) & brain_actual_vs_AOP_reports(ii, 3)
        End If
        
        'cola os dados no report '* <----------------------
        Set rs = CurrentDb.OpenRecordset(sqlReport)
        If rs.RecordCount = 0 Then
          MsgBox "Não existem dad1os na base de dados do relatório '" & nomeReport & "': " & dataFinal, vbCritical
          Exit Sub
        End If
        
        Call InstanciaExcel.insereTitulo(nomeReport, brain_actual_vs_AOP_reports(ii, 2), False)
        Call InstanciaExcel.ColaDadosRS(rs, nomeReport, brain_actual_vs_AOP_reports(ii, 5), brain_actual_vs_AOP_reports(ii, 6))  'cola os dados do recordset na planilha | formata o estilo da header aqui
        Call InstanciaExcel.converteRangeTabela(nomeReport, nomeReport & ii, True, brain_actual_vs_AOP_reports(ii, 5), brain_actual_vs_AOP_reports(ii, 6), False) 'converte o range colado em tabela
              
      Next ii
      
    '---------------------------------------------------------------------
    'monta os reports de outputs - LATAM e MPR
    ElseIf relatorios(k) = "Outputs LATAM" And outputsLatam = "Outputs LATAM" Then
            
      reportOutputs(1, 1) = "Month vs AOP Output" 'nome da aba
      reportOutputs(1, 2) = "month_vs_AOP_output" 'consulta de categoria
      reportOutputs(1, 3) = "month_vs_AOP_BU_output" 'consulta de BU
      reportOutputs(1, 4) = "month_vs_AOP_BU_output_TOTAL" 'consulta de mes
      reportOutputs(1, 5) = "month_vs_AOP_output_CR" 'consulta por Region/Category
      
      reportOutputs(2, 1) = "Month vs PY Output" 'nome da aba
      reportOutputs(2, 2) = "month_vs_PY_output" 'consulta de categoria
      reportOutputs(2, 3) = "month_vs_PY_BU_output" 'consulta de BU
      reportOutputs(2, 4) = "month_vs_PY_BU_output_TOTAL" 'consulta de mes
      reportOutputs(2, 5) = "month_vs_PY_output_CR" 'consulta por Region/Category
      
      reportOutputs(3, 1) = "YTD vs AOP Output" 'nome da aba
      reportOutputs(3, 2) = "month_vs_AOP_output_agrupado" 'consulta de categoria
      reportOutputs(3, 3) = "month_vs_AOP_BU_output_agrupado" 'consulta de BU
      reportOutputs(3, 4) = "month_vs_AOP_BU_output_TOTAL_agrupado" 'consulta de mes
      reportOutputs(3, 5) = "month_vs_AOP_output_CR_agrupado" 'consulta por Region/Category

      reportOutputs(4, 1) = "YTD vs PY Output" 'nome da aba
      reportOutputs(4, 2) = "month_vs_PY_output_agrupado" 'consulta de categoria
      reportOutputs(4, 3) = "month_vs_PY_BU_output_agrupado" 'consulta de BU
      reportOutputs(4, 4) = "month_vs_PY_BU_output_TOTAL_agrupado" 'consulta de mes
      reportOutputs(4, 5) = "month_vs_PY_output_CR_agrupado" 'consulta por Region/Category

      'itera pelas abas, criando
      For j = 1 To UBound(reportOutputs)
                   
        'so adiciona sheets que forem novas e nao estiverem preenchidas
        If InstanciaExcel.xls.Sheets(1).UsedRange.Rows.Count > 1 Then
          InstanciaExcel.xls.Sheets.Add After:=InstanciaExcel.xls.Sheets(InstanciaExcel.xls.Sheets.Count)  'adiciona um sheet a cada iteracao
        End If
        
        '--------------------------------------
        'gerar os LATAM OUTPUTS

        'cria as sheets vazias
        nomeReport = reportOutputs(j, 1)
        InstanciaExcel.xls.Sheets(InstanciaExcel.xls.Sheets.Count).name = reportOutputs(j, 1)
  
        '---------------------------------------------------------------------
        'se for para gerar apenas uma BU
        sSQLdistinctCategoria = "select distinct filtraBU.BU from filtraBU inner join mpr_month_vs_AOP_union on filtraBU.BU = mpr_month_vs_AOP_union.BU order by filtraBU.BU asc"
'        If chk_BU Then
'          sSQLdistinctCategoria = " select distinct " & _
'                                  " BU " & _
'                                  " from " & _
'                                  "   categoryMIX_report_final_AOP_bd_a " & _
'                                  " where BU='" & cbo_BU & "' " & _
'                                  " order by " & _
'                                  "   BU asc "
'        Else
'          sSQLdistinctCategoria = " select distinct " & _
'                                  " BU " & _
'                                  " from " & _
'                                  "   categoryMIX_report_final_AOP_bd_a " & _
'                                  " order by " & _
'                                  "   BU asc "
'        End If
        '---------------------------------------------------------------------
        Set rsDistinct = CurrentDb.OpenRecordset(sSQLdistinctCategoria)
        
        '0 iteracao colar o total
        
        'cola Total
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        'se for actual sao os 2 primeiros itens do array ||| os ACTUAL sao os 3 e 4 do array
        If j <= 2 Then
          Set rsTotal = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 4) & " where Category<>'0' and Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# ")
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        'os reports de YTD nao tem o month como criterio ||| os YTD sao os 3 e 4 do array
        Else
          Set rsTotal = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 4))
        End If
        
        Call InstanciaExcel.insereTitulo(nomeReport, "TOTAL", False) 'titulo
        Call InstanciaExcel.ColaDadosRS(rsTotal, nomeReport, "B", 2)  'cola os dados do recordset na planilha | formata o estilo da header aqui
        Call InstanciaExcel.converteRangeTabela(nomeReport, reportOutputs(j, 1) & j, True, "B", 2, True)  'converte o range colado em tabela
        
        'faz a iteraçao de colagem de dados
        While Not rsDistinct.EOF
          
          '--------------------------------------------------
          '1 iteracao colar o total de BU [BU]
          '2 iteracao colar cada [Region/Channel/Categoria-BU]
          
          '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
          'se for actual sao os 2 primeiros itens do array ||| os ACTUAL sao os 3 e 4 do array
          If j <= 2 Then
            Set rsBU = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 3) & " where Category<>'0' and BU='" & rsDistinct.Fields("BU").Value & "' and Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# ")
            Set rsRegionChannel = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 2) & " where Category<>'0' and BU='" & rsDistinct.Fields("BU").Value & "' and Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# ")
            Set rsRegionsChannelReal = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 5) & " where Category<>'0' and BU='" & rsDistinct.Fields("BU").Value & "' and Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# order by BU asc, Month asc, [Region/Channel] asc, Category asc")
            
          '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
          'os reports de YTD nao tem o month como criterio ||| os YTD sao os 3 e 4 do array
          Else
            Set rsBU = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 3) & " where BU='" & rsDistinct.Fields("BU").Value & "'")
            Set rsRegionChannel = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 2) & " where BU='" & rsDistinct.Fields("BU").Value & "'")
            Set rsRegionsChannelReal = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 5) & " where BU='" & rsDistinct.Fields("BU").Value & "' order by BU asc, [Region/Channel] asc, Category asc ")
          End If
          '--------------------------------------------------
 
          Call InstanciaExcel.insereTitulo(nomeReport, rsDistinct.Fields("BU").Value, True) 'titulo
          
          'cola BU
          Call InstanciaExcel.insereTitulo(nomeReport, "BU", False) 'titulo
          Call InstanciaExcel.ColaDadosRS(rsBU, nomeReport, "B", 2)  'cola os dados do recordset na planilha | formata o estilo da header aqui
          Call InstanciaExcel.converteRangeTabela(nomeReport, reportOutputs(j, 1) & j, True, "B", 2, False)  'converte o range colado em tabela

          'cola Categoria
          Call InstanciaExcel.insereTitulo(nomeReport, "Categoria", False)
          Call InstanciaExcel.ColaDadosRS(rsRegionChannel, nomeReport, "B", 2)  'cola os dados do recordset na planilha | formata o estilo da header aqui
          Call InstanciaExcel.converteRangeTabela(nomeReport, reportOutputs(j, 1) & j, True, "B", "2", False) 'converte o range colado em tabela
         
          'cola Categoria/Region
          Call InstanciaExcel.insereTitulo(nomeReport, "Categoria/Region", False)
          Call InstanciaExcel.ColaDadosRS(rsRegionsChannelReal, nomeReport, "B", 2)  'cola os dados do recordset na planilha | formata o estilo da header aqui
          Call InstanciaExcel.converteRangeTabela(nomeReport, reportOutputs(j, 1) & j, True, "B", "2", False) 'converte o range colado em tabela
                    
          rsDistinct.MoveNext 'vai para o proximo
        Wend 'proxima BU (BU/Categoria)
        
      Next j 'proxima aba (...actual aop/py, ytd aop/py)
        
    '---------------------------------------------------------------------
    'gera os reports de MPR
    ElseIf relatorios(k) = "Outputs MPR" And outputsLatam = "Outputs MPR" Then
            
      reportOutputs(1, 1) = "MPR Month vs AOP" 'nome da aba
      reportOutputs(1, 2) = "mpr_month_vs_AOP_union" 'consulta de categoria
      reportOutputs(1, 3) = "mpr_month_vs_AOP_BU" 'consulta de BU
      reportOutputs(1, 4) = "mpr_month_vs_AOP_BU_total" 'consulta de mes
      
      reportOutputs(2, 1) = "MPR Month vs PY" 'nome da aba
      reportOutputs(2, 2) = "mpr_month_vs_PY_union" 'consulta de categoria
      reportOutputs(2, 3) = "mpr_month_vs_PY_BU" 'consulta de BU
      reportOutputs(2, 4) = "mpr_month_vs_PY_BU_total" 'consulta de mes
      
      reportOutputs(3, 1) = "MPR YTD vs AOP" 'nome da aba
      reportOutputs(3, 2) = "mpr_month_vs_AOP_agrupado" 'consulta de categoria
      reportOutputs(3, 3) = "mpr_month_vs_AOP_BU_agrupado" 'consulta de BU
      reportOutputs(3, 4) = "mpr_month_vs_AOP_BU_total_agrupado" 'consulta de mes
      
      reportOutputs(4, 1) = "MPR YTD vs PY" 'nome da aba
      reportOutputs(4, 2) = "mpr_month_vs_PY_agrupado" 'consulta de categoria
      reportOutputs(4, 3) = "mpr_month_vs_PY_BU_agrupado" 'consulta de BU
      reportOutputs(4, 4) = "mpr_month_vs_PY_BU_total_agrupado" 'consulta de mes
      
      
      'itera pelas abas, criando
      For j = 1 To UBound(reportOutputs)
                   
        'so adiciona sheets que forem novas e nao estiverem preenchidas
        If InstanciaExcel.xls.Sheets(1).UsedRange.Rows.Count > 1 Then
          InstanciaExcel.xls.Sheets.Add After:=InstanciaExcel.xls.Sheets(InstanciaExcel.xls.Sheets.Count)  'adiciona um sheet a cada iteracao
        End If
        
        '--------------------------------------
        'gerar os MPR OUTPUTS
        'cria as sheets vazias
        nomeReport = reportOutputs(j, 1)
        InstanciaExcel.xls.Sheets(InstanciaExcel.xls.Sheets.Count).name = reportOutputs(j, 1)
        
        '---------------------------------------------------------------------
        'se for para gerar apenas uma BU
        sSQLdistinctCategoria = "select distinct filtraBU.BU from filtraBU inner join mpr_month_vs_AOP_union on filtraBU.BU = mpr_month_vs_AOP_union.BU order by filtraBU.BU asc"

'
'        If chk_BU Then
'          sSQLdistinctCategoria = " select distinct " & _
'                                  " BU " & _
'                                  " from " & _
'                                  "   categoryMIX_report_final_AOP_bd_a " & _
'                                  " where BU='" & cbo_BU & "' " & _
'                                  " order by " & _
'                                  "   BU asc "
'        Else
'          sSQLdistinctCategoria = " select distinct " & _
'                                  " BU " & _
'                                  " from " & _
'                                  "   categoryMIX_report_final_AOP_bd_a " & _
'                                  " order by " & _
'                                  "   BU asc "
'        End If
        '---------------------------------------------------------------------
        
        Set rsDistinct = CurrentDb.OpenRecordset(sSQLdistinctCategoria)
        
        '0 iteracao colar o total
        
        'cola Total
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        'se for actual sao os 2 primeiros itens do array ||| os ACTUAL sao os 3 e 4 do array
        If j <= 2 Then
          Set rsTotal = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 4) & " where Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# ")
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        'os reports de YTD nao tem o month como criterio ||| os YTD sao os 3 e 4 do array
        Else
          Set rsTotal = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 4))
        End If
        
        Call InstanciaExcel.insereTitulo(nomeReport, "TOTAL", False) 'titulo
        Call InstanciaExcel.ColaDadosRS(rsTotal, nomeReport, "B", 2)  'cola os dados do recordset na planilha | formata o estilo da header aqui
        Call InstanciaExcel.converteRangeTabela(nomeReport, reportOutputs(j, 1) & j, True, "B", 2, True)  'converte o range colado em tabela
        
        'faz a iteraçao de colagem de dados
        While Not rsDistinct.EOF
          
          '--------------------------------------------------
          '1 iteracao colar o total de BU [BU]
          '2 iteracao colar cada [Region/Channel/Categoria-BU]
          
          '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
          'se for actual sao os 2 primeiros itens do array ||| os ACTUAL sao os 3 e 4 do array
          If j <= 2 Then
            Set rsBU = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 3) & " where Category<>'0' and BU='" & rsDistinct.Fields("BU").Value & "' and Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# ")
            Set rsRegionChannel = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 2) & " where Category<>'0' and  BU='" & rsDistinct.Fields("BU").Value & "' and Month between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# ")
            
          '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
          'os reports de YTD nao tem o month como criterio ||| os YTD sao os 3 e 4 do array
          Else
            Set rsBU = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 3) & " where Category<>'0' and BU='" & rsDistinct.Fields("BU").Value & "'")
            Set rsRegionChannel = CurrentDb.OpenRecordset("select * from " & reportOutputs(j, 2) & " where Category<>'0' and BU='" & rsDistinct.Fields("BU").Value & "'")
          End If
          '--------------------------------------------------
 
          Call InstanciaExcel.insereTitulo(nomeReport, rsDistinct.Fields("BU").Value, True) 'titulo
          
          'cola BU
          Call InstanciaExcel.insereTitulo(nomeReport, "BU", False) 'titulo
          Call InstanciaExcel.ColaDadosRS(rsBU, nomeReport, "B", 2)  'cola os dados do recordset na planilha | formata o estilo da header aqui
          Call InstanciaExcel.converteRangeTabela(nomeReport, reportOutputs(j, 1) & j, True, "B", 2, False)  'converte o range colado em tabela

          'cola Categoria/Region
          Call InstanciaExcel.insereTitulo(nomeReport, "Categoria/Region", False)
          Call InstanciaExcel.ColaDadosRS(rsRegionChannel, nomeReport, "B", 2)  'cola os dados do recordset na planilha | formata o estilo da header aqui
          Call InstanciaExcel.converteRangeTabela(nomeReport, reportOutputs(j, 1) & j, True, "B", "2", False) 'converte o range colado em tabela
         
          rsDistinct.MoveNext 'vai para o proximo
        Wend 'proxima BU (BU/Categoria)
        
      Next j 'proxima aba (...actual aop/py, ytd aop/py)
    
  
    '---------------------------------------------------------------------------------------
    'SCORECARDS
    'gera os reports de MPR
    ElseIf relatorios(k) = "Scorecards" And outputsLatam = "Scorecards" Then
      
      
      '**************************************************************
      'NSV
      reportScorecards(1, 1) = "scorecard_NSV_vs_AOP_TOTAL" 'total
      reportScorecards(1, 2) = "scorecard_NSV_vs_AOP" 'detalhe
      reportScorecards(1, 3) = False 'is agrupado (YTD)
      reportScorecards(1, 4) = "Month Zone Scorecard vs. Budget" 'titulo da coluna
      reportScorecards(1, 5) = "B" 'coluna inicial
      reportScorecards(1, 6) = 2 'coluna inicial
      reportScorecards(1, 7) = "NSV" 'titulo da secao
       
      reportScorecards(2, 1) = "scorecard_NSV_vs_PY_TOTAL" 'total
      reportScorecards(2, 2) = "scorecard_NSV_vs_PY" 'detalhe
      reportScorecards(2, 3) = False 'is agrupado (YTD)
      reportScorecards(2, 4) = "Month Zone Scorecard vs. PY" 'titulo da coluna
      reportScorecards(2, 5) = "L" 'coluna inicial
      reportScorecards(2, 6) = 12 'coluna inicial
      reportScorecards(2, 7) = "NSV" 'titulo da secao
      
      reportScorecards(3, 1) = "scorecard_NSV_vs_AOP_agrupado_TOTAL" 'total
      reportScorecards(3, 2) = "scorecard_NSV_vs_AOP_agrupado" 'detalhe
      reportScorecards(3, 3) = True 'is agrupado (YTD)
      reportScorecards(3, 4) = "YTD Zone Scorecard vs. AOP" 'titulo da coluna
      reportScorecards(3, 5) = "V" 'coluna inicial
      reportScorecards(3, 6) = 22 'coluna inicial
      reportScorecards(3, 7) = "NSV" 'titulo da secao
      
      reportScorecards(4, 1) = "scorecard_NSV_vs_PY_agrupado_TOTAL" 'total
      reportScorecards(4, 2) = "scorecard_NSV_vs_PY_agrupado" 'detalhe
      reportScorecards(4, 3) = True 'is agrupado (YTD)
      reportScorecards(4, 4) = "YTD Zone Scorecard vs. PY" 'titulo da coluna
      reportScorecards(4, 5) = "AF" 'coluna inicial
      reportScorecards(4, 6) = 32 'coluna inicial
      reportScorecards(4, 7) = "NSV" 'titulo da secao
      
      '**************************************************************
      'CMA
      reportScorecards(5, 1) = "scorecard_CMA_vs_AOP_TOTAL" 'total
      reportScorecards(5, 2) = "scorecard_CMA_vs_AOP" 'detalhe
      reportScorecards(5, 3) = False 'is agrupado (YTD)
      reportScorecards(5, 5) = "B" 'coluna inicial
      reportScorecards(5, 6) = 2 'coluna inicial
      reportScorecards(5, 7) = "CMA" 'titulo da secao
       
      reportScorecards(6, 1) = "scorecard_CMA_vs_PY_TOTAL" 'total
      reportScorecards(6, 2) = "scorecard_CMA_vs_PY" 'detalhe
      reportScorecards(6, 3) = False 'is agrupado (YTD)
      reportScorecards(6, 5) = "L" 'coluna inicial
      reportScorecards(6, 6) = 12 'coluna inicial
      reportScorecards(6, 7) = "CMA" 'titulo da secao
      
      reportScorecards(7, 1) = "scorecard_CMA_vs_AOP_agrupado_TOTAL" 'total
      reportScorecards(7, 2) = "scorecard_CMA_vs_AOP_agrupado" 'detalhe
      reportScorecards(7, 3) = True 'is agrupado (YTD)
      reportScorecards(7, 5) = "V" 'coluna inicial
      reportScorecards(7, 6) = 22 'coluna inicial
      reportScorecards(7, 7) = "CMA" 'titulo da secao
      
      reportScorecards(8, 1) = "scorecard_CMA_vs_PY_agrupado_TOTAL" 'total
      reportScorecards(8, 2) = "scorecard_CMA_vs_PY_agrupado" 'detalhe
      reportScorecards(8, 3) = True 'is agrupado (YTD)
      reportScorecards(8, 5) = "AF" 'coluna inicial
      reportScorecards(8, 6) = 32 'coluna inicial
      reportScorecards(8, 7) = "CMA" 'titulo da secao
      
      '**************************************************************
      'CCMA
      '(ACTUAL)
      reportScorecards(9, 1) = "scorecard_UNION_AOP" 'detalhe
      reportScorecards(9, 2) = "" 'detalhe
      reportScorecards(9, 3) = False 'is agrupado (YTD)
      reportScorecards(9, 5) = "B" 'coluna inicial
      reportScorecards(9, 6) = 2 'coluna inicial
      reportScorecards(9, 7) = "EBITDA" 'titulo da secao
      
      reportScorecards(10, 1) = "scorecard_UNION_PY" 'detalhe
      reportScorecards(10, 2) = "" 'detalhe
      reportScorecards(10, 3) = False 'is agrupado (YTD)
      reportScorecards(10, 5) = "L" 'coluna inicial
      reportScorecards(10, 6) = 12 'coluna inicial
      reportScorecards(10, 7) = "EBITDA" 'titulo da secao
      
      '(YTD)
      reportScorecards(11, 1) = "scorecard_UNION_agrupado_AOP" 'detalhe
      reportScorecards(11, 2) = "" 'detalhe
      reportScorecards(11, 3) = True 'is agrupado (YTD)
      reportScorecards(11, 5) = "V" 'coluna inicial
      reportScorecards(11, 6) = 22 'coluna inicial
      reportScorecards(11, 7) = "EBITDA" 'titulo da secao
      
      reportScorecards(12, 1) = "scorecard_UNION_agrupado_PY" 'detalhe
      reportScorecards(12, 2) = "" 'detalhe
      reportScorecards(12, 3) = True 'is agrupado (YTD)
      reportScorecards(12, 5) = "AF" 'coluna inicial
      reportScorecards(12, 6) = 32 'coluna inicial
      reportScorecards(12, 7) = "EBITDA" 'titulo da secao


      '-------------------------------------------------------------------
      'so adiciona 1 sheet
      If InstanciaExcel.xls.Sheets(1).UsedRange.Rows.Count > 1 Then
        InstanciaExcel.xls.Sheets.Add After:=InstanciaExcel.xls.Sheets(InstanciaExcel.xls.Sheets.Count)
      End If
      
      nomeReport = "Latam Scorecards"
      InstanciaExcel.xls.Sheets(InstanciaExcel.xls.Sheets.Count).name = nomeReport
      
      '-------------------------------------------------------------------
      'itera pelos scorecards, || tem que colar coluna lado-a-lado
      For y = 1 To UBound(reportScorecards)
      
        '**************************
        'coloca os titulos na colunas
        If reportScorecards(y, 4) <> "" Then
          Call InstanciaExcel.insereTituloScorecards(nomeReport, reportScorecards(y, 4), False, reportScorecards(y, 5), reportScorecards(y, 6))   'titulo da coluna
        End If
                
        'se for Actual
        If reportScorecards(y, 3) = False Then
          Set rsTotal = CurrentDb.OpenRecordset("select * from " & reportScorecards(y, 1) & " where Month between #" & Format(dataFinal, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# ") 'total
          
          If reportScorecards(y, 2) <> "" Then
            Set rsDetalhe = CurrentDb.OpenRecordset("select * from " & reportScorecards(y, 2) & " where Month between #" & Format(dataFinal, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# ") 'detalhe
          End If
          
        'se for YTD
        ElseIf reportScorecards(y, 3) = True Then
          Set rsTotal = CurrentDb.OpenRecordset("select * from " & reportScorecards(y, 1)) 'total
          
          If reportScorecards(y, 2) <> "" Then
            Set rsDetalhe = CurrentDb.OpenRecordset("select * from " & reportScorecards(y, 2)) 'detalhe
          End If
        End If
        
        'titulo da secao
        Call InstanciaExcel.insereTituloScorecards(nomeReport, reportScorecards(y, 7), True, reportScorecards(y, 5), reportScorecards(y, 6))    'titulo da coluna
        
        '**********************************
        'Cola o total
        'Call InstanciaExcel.insereTitulo(nomeReport, "TOTAL", False) 'titulo
        Call InstanciaExcel.ColaDadosRS_Scorecards(rsTotal, nomeReport, reportScorecards(y, 5), reportScorecards(y, 6))  'cola os dados do recordset na planilha | formata o estilo da header aqui
        Call InstanciaExcel.converteRangeTabela_Scorecards(nomeReport, reportScorecards(y, 1) & y, True, reportScorecards(y, 5), reportScorecards(y, 6), True)  'converte o range colado em tabela
        
        '**********************************
        'Cola o detalhe
        If reportScorecards(y, 2) <> "" Then
          Call InstanciaExcel.ColaDadosRS_Scorecards(rsDetalhe, nomeReport, reportScorecards(y, 5), reportScorecards(y, 6))  'cola os dados do recordset na planilha | formata o estilo da header aqui
          Call InstanciaExcel.converteRangeTabela_Scorecards(nomeReport, reportScorecards(y, 1) & y, True, reportScorecards(y, 5), reportScorecards(y, 6), True)  'converte o range colado em tabela
        End If
      Next y

    End If

nextIteration:
  Next k

End Sub

'habilita os controles de geração de formulário
Public Sub habilitaAgrupado(ByRef frm_gerarReport As Form)
  
  'habilita as datas
  If frm_gerarReport.relatorioAgrupado.Value Then
    frm_gerarReport.dataInicial.Enabled = True
  
  'desabilita as datas
  ElseIf frm_gerarReport.relatorioAgrupado.Value = False And (itemSelecionado(frm_gerarReport.lista_reports, "Outputs LATAM") <> 1 And itemSelecionado(frm_gerarReport.lista_reports, "Outputs MPR") <> 1) Then
    frm_gerarReport.dataInicial.Enabled = False
    frm_gerarReport.dataInicial.Value = ""
  End If
  
End Sub

'formata os reports
Public Sub formataReports(ByRef xls As Excel.Application, reportType As String, ByRef InstanciaExcel As InstanciaExcel, dataFinal As String)

  Dim tabela As ListObject
  Dim cell As Range
  Dim comecoColuna As String
  Dim identificaTons As Boolean
  Dim mudaCor As Boolean
  Dim aba As Worksheet
  Dim cabecalho As Range
  Dim totalColunasCinzas As Integer
  
  'itera pelas sheets
  For Each aba In xls.Sheets
    
    'aumenta a largura de todas as colunas para 18cm
    aba.Select
           
    '-------------------------------------------------
    'itera pelas tabelas
    For Each tabela In aba.ListObjects
      
      xls.Cells.ColumnWidth = 18 'altera largura das colunas
      
      Set cabecalho = tabela.HeaderRowRange
      
      tabela.TableStyle = ""
      tabela.Unlist 'remove a tabela da sheet
      
      cabecalho.Interior.Color = 11892015 'muda as cores pra azul
      cabecalho.Font.Color = rgbWhite
      
      '--------------------------------------------------------------------------
      'pinta a parte superior de cinza
      If aba.name <> "INPUT Scenario" Then
        
        'se nao for 'Output'
        If reportType = "Output" Then
          totalColunasCinzas = 4
          
        '-----------------------------------------------
        ElseIf reportType = "MPR" Then
          totalColunasCinzas = 3
          
        '-----------------------------------------------
        ElseIf reportType = "Scorecards" Then
          totalColunasCinzas = 3
          xls.Cells.ColumnWidth = 9 'altera largura das colunas
          
        '-----------------------------------------------
        'se nao for 'Output' //// ->Brain
        Else
          'numero fixado de colunas
          If cabecalho.Columns.Count = 51 Then
            totalColunasCinzas = 5
          Else
            totalColunasCinzas = 6
          End If
        End If
        
        cabecalho.Offset(-1, totalColunasCinzas).Resize(cabecalho.Rows.Count, cabecalho.Columns.Count - totalColunasCinzas).Interior.Color = 10921638 'cinza
        cabecalho.Offset(-1, totalColunasCinzas).Resize(cabecalho.Rows.Count, cabecalho.Columns.Count - totalColunasCinzas).Font.Color = rgbWhite
        Call alteraBorda(cabecalho.Offset(-1, totalColunasCinzas).Resize(cabecalho.Rows.Count, cabecalho.Columns.Count - totalColunasCinzas)) 'altera a borda da celula para branco

        'itera pela header row
        For Each cell In cabecalho
          '--------------------------------------------------------------------------
          'identifica a parte que tem "nome|" e coloca em cima **Pinta de cinza
          If cell.Value Like "*|*" Then
            cell.Offset(-1, 0).Value = Left(cell.Value, InStr(1, cell.Value, "|") - 1) 'titulo de cima
            cell.Value = Mid(cell.Value, InStr(1, cell.Value, "|") + 1, 300) 'titulo de baixo
          End If
          
          '--------------------------------------------------------------------------
          'identifica a parte que tem "_" e substitui por nada
          If cell.Value Like "*_*" Then
            cell.Value = Replace(cell.Value, "_", "") 'substitui por nada
          End If
          
          '--------------------------------------------------------------------------
          'substitui 'year' por ano
          If cell.Value Like "*year*" Then
            cell.Value = Replace(cell.Value, "year", "'" & Year(CDate(dataFinal))) 'titulo de baixo
          End If
        Next cell
      End If '//If aba.name <> "INPUT Scenario"
      
    Next tabela
    
    '-------------------------------------------------
    'formata colunas
    aba.UsedRange.Font.Size = 9
    aba.UsedRange.RowHeight = 12
    aba.Columns("A:A").ColumnWidth = 3 'coluna a estreita
    
    '-------------------------------------------------
    If aba.name <> "INPUT Scenario" Then
      aba.Columns("A:A").Interior.Color = rgbWhite
      aba.Columns("A:A").Font.Color = rgbWhite
      aba.Columns("A:A").NumberFormat = "#,##0"
            
      '<--------------------------------
      'Se for uma SCORECARDS '*********very hardcoded
      If reportType = "Scorecards" And aba.name Like "*Scorecards*" Then
        Call criaColunasVaziasScorecards(aba)  'cria as colunas // para os Outputs
        Call InstanciaExcel.formataCasasDecimaisDaSheet(aba.name) 'campos com 3 casas decimais
        Call formatDecimaisScorecards(aba)
        
      '<--------------------------------
      'Se for uma OUTPUT '*********very hardcoded
      ElseIf reportType = "Output" And aba.name Like "*Output*" Then
        Call criaColunasVaziasOutputs(aba)  'cria as colunas // para os Outputs
        Call InstanciaExcel.formataCasasDecimaisDaSheet(aba.name) 'campos com 3 casas decimais
        Call formatDecimaisOutputs(aba)

      '<--------------------------------
      'Se for uma MPR '*********very hardcoded
      ElseIf reportType = "MPR" And aba.name Like "*MPR*" Then
        Call criaColunasVaziasMPR(aba)  'cria as colunas vazias apos a geracao
        Call InstanciaExcel.formataCasasDecimaisDaSheet(aba.name)
        Call formatDecimaisMPR(aba)

      '<--------------------------------
      'Se for uma brain normal '*********very hardcoded
      ElseIf reportType = "Brain" Then
        Call criaColunasVazias(aba)  'cria as colunas vazias apos a geracao
        Call InstanciaExcel.formataCasasDecimaisDaSheet(aba.name)
      End If
    End If
    
    Call centralizaCelula(aba.UsedRange) 'centraliza a celula
    
    aba.Columns("C:C").NumberFormat = "mmm-yyyy" 'formata data
    xls.ActiveWindow.DisplayGridlines = False 'remove as gridlines
  Next aba
  
  
End Sub

''itera pelos chars da string
'Private Function primeiroUnderline(palavra As String) As Integer
'  Dim i As Integer
'  For i = Len(palavra) To 1 Step -1
'    ultimoUnderline = i
'    Exit Function
'  Next i
'End Function

Private Function bol(b As Boolean) As Boolean
  If b = True Then
    bol = False
  Else
    bol = True
  End If
End Function

'pega a dataInicial e dataFinal e insere na filtro agrupado
Public Sub insereNaFiltroAgrupado(dataInicio As String, dataFinal As String)
  
  Dim rs As Recordset
  Dim contagemIniciada As Boolean
  Dim datasArray() As String
  Dim k As Long
  Dim rsfiltraAgrupado As Recordset
  Dim valor As Variant
  Set rsfiltraAgrupado = CurrentDb.OpenRecordset("select * from filtraAgrupado")
  
  Set rs = CurrentDb.OpenRecordset("select * from datas")

  While Not rs.EOF
    '----------------------------
    If rs.Fields("datas") = dataInicio Then contagemIniciada = True
    '----------------------------
    If contagemIniciada Then
        k = k + 1
      ReDim Preserve datasArray(1 To k)
      datasArray(k) = rs.Fields("datas").Value
    End If
    
    '----------------------------
    If rs.Fields("datas") = dataFinal Then
      
      DoCmd.RunSQL "delete * from filtraAgrupado"
      
      'insere os registros na 'filtraAgrupado'
      For Each valor In datasArray
        rsfiltraAgrupado.AddNew 'cria um novo registro
        rsfiltraAgrupado.Fields("dataFiltro").Value = CDate(valor)
        rsfiltraAgrupado.Update 'insere um novo registro
      Next valor
      Exit Sub
    End If
    '----------------------------
    rs.MoveNext
  Wend
  
End Sub

'insere filtro de BU
Public Sub insereFiltroBU(BU As String)
  DoCmd.RunSQL "delete * from filtraBU"
  DoCmd.RunSQL "insert into filtraBU(BU) values('" & BU & "')"
End Sub

'insere filtro de BU
Public Sub nenhumNaFiltroBU()

  Dim sSQL As String
  DoCmd.RunSQL "delete * from filtraBU"
  
  '-------------------------------------------------------------------
  'se nao houver filtro na BU, da o insert de todas as BU
         sSQL = " insert into  "
  sSQL = sSQL & "     filtraBU (BU) "
  sSQL = sSQL & " select distinct  "
  sSQL = sSQL & "   base_report_INFERIOR_t_bd.BU "
  sSQL = sSQL & " from  "
  sSQL = sSQL & "   base_report_INFERIOR_t_bd left join filtraBU on "
  sSQL = sSQL & "   base_report_INFERIOR_t_bd.BU = filtraBU.BU "
  sSQL = sSQL & " where "
  sSQL = sSQL & "   filtraBU.BU is null "
    
  DoCmd.RunSQL sSQL
End Sub

'cria as colunas vazias apos a geracao
Public Sub formatDecimaisOutputs(ByRef aba As Worksheet)
  
  Dim arrColunas(1 To 16)
  Dim aspaDuplo As String
  Dim joinsJuntos As String
  aspaDuplo = """"
  Dim v As Variant
  
  arrColunas(1) = "J"
  arrColunas(2) = "L"
  arrColunas(3) = "M"
  arrColunas(4) = "R"
  arrColunas(5) = "W"
  arrColunas(6) = "AB"
  arrColunas(7) = "AG"
  arrColunas(8) = "AL"
  arrColunas(9) = "AQ"
  arrColunas(10) = "AV"
  arrColunas(11) = "BA"
  arrColunas(12) = "BF"
  arrColunas(13) = "BK"
  arrColunas(14) = "BM"
  arrColunas(15) = "BN"
  arrColunas(16) = "BO"
  
  For Each v In arrColunas
    'Columns(v & ":" & v).Select
    aba.Columns(v & ":" & v).NumberFormat = "0.00%"
    'aba.Columns(v & ":" & v).ColumnWidth = 1.5
    'aba.Columns(v & ":" & v).Clear
  Next v
  
End Sub

'cria as colunas vazias apos a geracao
Public Sub formatDecimaisScorecards(ByRef aba As Worksheet)
  
  Dim arrColunas(1 To 4)
  Dim aspaDuplo As String
  Dim joinsJuntos As String
  aspaDuplo = """"
  Dim v As Variant
  
  arrColunas(1) = "I"
  arrColunas(2) = "T"
  arrColunas(3) = "AE"
  arrColunas(4) = "AL"
  
  For Each v In arrColunas
    'Columns(v & ":" & v).Select
    aba.Columns(v & ":" & v).NumberFormat = "0.00%"
    'aba.Columns(v & ":" & v).ColumnWidth = 1.5
    'aba.Columns(v & ":" & v).Clear
  Next v
  
  '======================================================
  Dim colDatas(1 To 1) As String
  colDatas(1) = "N"
  
  For Each v In colDatas
    'Columns(v & ":" & v).Select
    aba.Columns(v & ":" & v).NumberFormat = "mmm-yyyy"
    'aba.Columns(v & ":" & v).ColumnWidth = 1.5
    'aba.Columns(v & ":" & v).Clear
  Next v
  
End Sub

'cria as colunas vazias apos a geracao
Public Sub formatDecimaisMPR(ByRef aba As Worksheet)
  
  Dim arrColunas(1 To 5)
  Dim aspaDuplo As String
  Dim joinsJuntos As String
  aspaDuplo = """"
  Dim v As Variant
  
  arrColunas(1) = "G"
  arrColunas(2) = "J"
  arrColunas(3) = "S"
  arrColunas(4) = "V"
  arrColunas(5) = "W"
  
  For Each v In arrColunas
    'Columns(v & ":" & v).Select
    aba.Columns(v & ":" & v).NumberFormat = "0.00%"
    'aba.Columns(v & ":" & v).ColumnWidth = 1.5
    'aba.Columns(v & ":" & v).Clear
  Next v
  
End Sub

'cria as colunas vazias apos a geracao
Public Sub criaColunasVaziasScorecards(ByRef aba As Worksheet)
  
  Dim arrColunas(1 To 4)
  Dim aspaDuplo As String
  Dim joinsJuntos As String
  aspaDuplo = """"
  Dim v As Variant
  
  arrColunas(1) = "E"
  arrColunas(2) = "P"
  arrColunas(3) = "AA"
  arrColunas(4) = "AL"

  For Each v In arrColunas
    'Columns(v & ":" & v).Select
    aba.Columns(v & ":" & v).Insert
    aba.Columns(v & ":" & v).ColumnWidth = 0.7
    aba.Columns(v & ":" & v).Clear
  Next v
  
End Sub

'cria as colunas vazias apos a geracao
Public Sub criaColunasVaziasOutputs(ByRef aba As Worksheet)
  
  Dim arrColunas(1 To 18)
  Dim aspaDuplo As String
  Dim joinsJuntos As String
  aspaDuplo = """"
  Dim v As Variant
  
  arrColunas(1) = "F"
  arrColunas(2) = "K"
  arrColunas(3) = "N"
  arrColunas(4) = "S"
  arrColunas(5) = "X"
  arrColunas(6) = "AC"
  arrColunas(7) = "AH"
  arrColunas(8) = "AM"
  arrColunas(9) = "AR"
  arrColunas(10) = "AW"
  arrColunas(11) = "BB"
  arrColunas(12) = "BG"
  arrColunas(13) = "BL"
  arrColunas(14) = "BP"
  arrColunas(15) = "BX"
  arrColunas(16) = "CE"
  arrColunas(17) = "CL"
  arrColunas(18) = "CU"
  
  For Each v In arrColunas
    'Columns(v & ":" & v).Select
    aba.Columns(v & ":" & v).Insert
    aba.Columns(v & ":" & v).ColumnWidth = 0.7
    aba.Columns(v & ":" & v).Clear
  Next v
  
End Sub

'cria as colunas vazias apos a geracao
Public Sub criaColunasVazias(ByRef aba As Worksheet)
  
  Dim arrColunas(1 To 10)
  Dim aspaDuplo As String
  Dim joinsJuntos As String
  aspaDuplo = """"
  Dim v As Variant
  
  arrColunas(1) = "G"
  arrColunas(2) = "M"
  arrColunas(3) = "R"
  arrColunas(4) = "X"
  arrColunas(5) = "AC"
  arrColunas(6) = "AJ"
  arrColunas(7) = "AQ"
  arrColunas(8) = "AV"
  arrColunas(9) = "BA"
  arrColunas(10) = "BF"
  
  For Each v In arrColunas
    'Columns(v & ":" & v).Select
    aba.Columns(v & ":" & v).Insert
    aba.Columns(v & ":" & v).ColumnWidth = 0.7
    aba.Columns(v & ":" & v).Clear
  Next v
  
End Sub

'cria as colunas vazias apos a geracao
Public Sub criaColunasVaziasMPR(ByRef aba As Worksheet)
  
  Dim arrColunas(1 To 7)
  Dim aspaDuplo As String
  Dim joinsJuntos As String
  aspaDuplo = """"
  Dim v As Variant
  
  arrColunas(1) = "E"
  arrColunas(2) = "H"
  arrColunas(3) = "K"
  arrColunas(4) = "N"
  arrColunas(5) = "Q"
  arrColunas(6) = "T"
  arrColunas(7) = "X"
  
  If aba.name Like "*MPR*" Then
  
    For Each v In arrColunas
      'Columns(v & ":" & v).Select
      aba.Columns(v & ":" & v).Insert
      aba.Columns(v & ":" & v).ColumnWidth = 0.7
      aba.Columns(v & ":" & v).Clear
    Next v
  End If
  
End Sub

'altera a borda da celula para branco
Private Sub alteraBorda(ByRef celula As Range)
    
  With celula.Borders(xlEdgeLeft)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = 0
      .Weight = xlThin
  End With
  
  With celula.Borders(xlEdgeTop)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = 0
      .Weight = xlThin
  End With
  
  With celula.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = 0
      .Weight = xlThin
  End With
  
  With celula.Borders(xlEdgeRight)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = 0
      .Weight = xlThin
  End With
    
    
  With celula.Borders(xlInsideVertical)
      .LineStyle = xlContinuous
      .ThemeColor = 1
      .TintAndShade = 0
      .Weight = xlThin
  End With
  celula.Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub

'centraliza o valor da celula
Private Sub centralizaCelula(ByRef celula As Range)
  celula.HorizontalAlignment = xlCenter 'centraliza
End Sub

'detecta se o item foi selecionado
Public Function itemSelecionado(ByRef lista_reports As ListBox, _
                                      reportTipo As String) As Integer

  Dim i As Integer

  '---------------------------------------
  For i = 0 To lista_reports.ListCount - 1
    
    If lista_reports.ItemData(i) Like "*" & reportTipo & "*" Then
      'se estiver selecionado
      If lista_reports.Selected(i) Then
        itemSelecionado = 1 'selected
        Exit Function
      Else
        itemSelecionado = 2 'unselected
      End If
    End If
  Next i
  
End Function

'detecta qual item foi selecionado
Public Function itemQualSelecionado(ByRef lista_reports As ListBox) As String

  Dim i As Integer
  Dim mesGringo As String

  '---------------------------------------
  For i = 0 To lista_reports.ListCount - 1
    
    'If lista_reports.ItemData(i) Like "*" & reportTipo & "*" Then
      'se estiver selecionado
      If lista_reports.Selected(i) Then
        mesGringo = Format(CDate(DLookup("datas", "datas", "data_formatada='" & lista_reports.Column(0) & "'")), "mm/dd/yyyy")
        itemQualSelecionado = " data=#" & mesGringo & "# and BU='" & lista_reports.Column(1) & "'"  'selected
        Exit Function
      End If
    'End If
  Next i
  
End Function


'detecta se o item foi selecionado
Public Function maisItensSelecionadosAlemDeOutputs(ByRef lista_reports As ListBox) As Integer

  Dim i As Integer

  '---------------------------------------
  For i = 0 To lista_reports.ListCount - 1
    
    If lista_reports.ItemData(i) <> "Outputs LATAM" And lista_reports.ItemData(i) <> "Outputs MPR" Then
      'se estiver selecionado
      If lista_reports.Selected(i) Then
        maisItensSelecionadosAlemDeOutputs = 1 'selected
        Exit Function
      End If
    End If
  Next i
  
End Function


'detecta se o item foi selecionado e habilita data inicial
Public Sub habilitaOutputs(ByRef frm_gerarReport As Form)

  'detecta se o item foi selecionado
  If (itemSelecionado(frm_gerarReport.lista_reports, "Outputs LATAM") = 1 Or itemSelecionado(frm_gerarReport.lista_reports, "Outputs MPR") = 1) Then
    frm_gerarReport.dataInicial.Enabled = True
    
  ElseIf (itemSelecionado(frm_gerarReport.lista_reports, "Outputs LATAM") = 2 Or itemSelecionado(frm_gerarReport.lista_reports, "Outputs MPR") = 2) And frm_gerarReport.relatorioAgrupado <> True Then
    frm_gerarReport.dataInicial.Enabled = False
    frm_gerarReport.dataInicial.Value = ""
  End If
End Sub

'Atualiza a combo de 'BU'
Public Sub atualiza_CBO_BU(ByRef frm_gerarReport As Form)
  
  Dim sSQL As String
  Dim i As ComboBox
  Dim dataInicio As String
  Dim dataFinal As String
  
   'Precisa selecionar as 2 datas - NUll
  If IsNull(frm_gerarReport.dataInicial.Value) Or IsNull(frm_gerarReport.dataFinal.Value) Then
    frm_gerarReport.cbo_BU.RowSource = "" 'deixa vazio
    Exit Sub
  End If
  
  'Precisa selecionar as 2 datas - Vazio
  If frm_gerarReport.dataInicial.Value = "" Or frm_gerarReport.dataFinal.Value = "" Then
    frm_gerarReport.cbo_BU.RowSource = "" 'deixa vazio
    Exit Sub
  End If
  
  dataInicio = frm_gerarReport.dataInicial.Value
  dataFinal = frm_gerarReport.dataFinal.Value
  'i.RowSource
  
         sSQL = " select distinct  "
  sSQL = sSQL & "   BU "
  sSQL = sSQL & " from "
  sSQL = sSQL & "  base_report_INFERIOR_t_bd "
  sSQL = sSQL & " where "
  sSQL = sSQL & "   data between #" & Format(dataInicio, "mm/dd/yyyy") & "# and #" & Format(dataFinal, "mm/dd/yyyy") & "# "

  frm_gerarReport.cbo_BU.Value = "" 'apaga o valor anterior
  frm_gerarReport.cbo_BU.RowSource = sSQL 'altera o rs da combo

End Sub

'habilita a combo de CBO_BU
Public Sub habilita_CBO_BU(ByRef frm_gerarReport As Form)

  If frm_gerarReport.chk_BU.Value Then
    frm_gerarReport.cbo_BU.Enabled = True
  Else
    frm_gerarReport.cbo_BU.Value = ""
    frm_gerarReport.cbo_BU.Enabled = False
  End If

End Sub
