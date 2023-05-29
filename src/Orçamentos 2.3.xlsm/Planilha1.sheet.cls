Private WithEvents qt As Excel.QueryTable

Sub verificaEstoque()

    Dim sqlGrupo, msgEstoque, qtdEstoque As String
    Dim rs As Recordset
    Dim demanda As Boolean
    
    demanda = True
    
    col = 2
    lin = 28
    
    sqlGrupo = "SELECT GW_CD_REF FROM PRODUTO WHERE PG_ID = 10000005"
    
        Do Until Cells(lin, col) = ""
            qtdEstoque = WorksheetFunction.IfError(Application.VLookup(Cells(lin, col), Sheets("Estoque").Range("A:F"), 6, False), "naoEncontrado")
            
            If qtdEstoque <> "naoEncontrado" Then
                Cells(lin, col).Select
                If ActiveCell.offset(0, 6).Value > CInt(qtdEstoque) Then
                
                    msgEstoque = "*Sem saldo de estoque disponível*"
                    
                    With Range(Cells(lin, col), ActiveCell.offset(0, 10))
                        .RowHeight = 45
                        .Interior.Color = 65535
                        .VerticalAlignment = xlCenter
                    End With
            
                    ActiveCell.offset(0, 1).Select
                                
                    Selection.WrapText = True
                    ActiveCell.Value = ActiveCell.Value & vbNewLine & msgEstoque
                                
                    With ActiveCell.Characters(Start:=Len(ActiveCell.Value) - Len(msgEstoque), Length:=Len(msgEstoque) + 1).Font
                        .name = "Calibri"
                        .FontStyle = "Negrito"
                        .Size = 8
                        .Color = -16776961
                    End With
                End If
            Else
                conexaoFireBird
                    Set rs = getRecordset(sqlGrupo)
                    Do While rs.EOF = False
                        If Cells(lin, col).Value = rs.Fields("GW_CD_REF") Then
                            demanda = False
                            Exit Do
                        Else
                            rs.MoveNext
                        End If
                    Loop
                fechaBancoDeDados
                
                If demanda = True Then
                    msgEstoque = "*Produto sob demanda*"
                    Cells(lin, col).Select
                    With Range(Cells(lin, col), ActiveCell.offset(0, 10))
                        .RowHeight = 45
                        .Interior.Color = 15773696
                        .VerticalAlignment = xlCenter
                    End With
            
                    ActiveCell.offset(0, 1).Select
                                
                    Selection.WrapText = True
                    ActiveCell.Value = ActiveCell.Value & vbNewLine & msgEstoque
                                
                    With ActiveCell.Characters(Start:=Len(ActiveCell.Value) - Len(msgEstoque), Length:=Len(msgEstoque) + 1).Font
                        .name = "Calibri"
                        .FontStyle = "Negrito"
                        .Size = 8
                        .Color = -16776961
                    End With
                End If
                
            End If
            lin = lin + 1
        Loop

End Sub

Function changeEnabled(ByVal bool As Boolean)

    Worksheets("Orçamento").Activate
    
    For controle = 1 To ActiveSheet.OLEObjects.Count
         If ActiveSheet.OLEObjects(controle).PrintObject = True Then ActiveSheet.OLEObjects(controle).Object.Enabled = bool
    Next
    
End Function

Sub relatorioDeTotais(ByVal valorSemDesc As Long)

    tbTotal.Value = Format(WorksheetFunction.Sum(Range("L:L")), "R$ ####.##")
    
    tbTotImp.Value = Format(WorksheetFunction.Sum(Range("L:L")) - WorksheetFunction.Sum(Range("K:K")), "R$ ####.##")
    
    tbTotDesc.Value = Format((valorSemDesc - WorksheetFunction.Sum(Range("L:L"))) + WorksheetFunction.Sum(Range("K:K")), "R$ ####.##")
    
    tbTotProd.Value = Format(valorSemDesc, "R$ ####.##")
    

End Sub

Sub buscarOrcamento()
    
    If Len(tbPainel.Value) < 5 Then
        MsgBox "O Nº Dav digitado não é válido, tente novamente!", VBA.VbMsgBoxStyle.vbInformation, "Gerenciamento de Dados"
    Else
        Application.ScreenUpdating = False
        
        Dim docVenda As Range
        Dim sqlProdutos, sqlCliente, sqlValTot As String
        Dim rs As Recordset
        Dim index As Integer
        Dim offset As Integer
        Dim valorTot As Long
        
        
        ActiveSheet.Unprotect
        changeEnabled (True)
        limparInformacoes
        
        sqlProdutos = "SELECT p.cd_ref, p.ds_completa, pvp.Qtd, pvp.Valor_Unitario, pvp.ValorLiquido," _
                    & "CASE WHEN p.nbm_sh = '32141020' THEN (pvp.ValorLiquido * 0.013) * pvp.Qtd ELSE 0 END AS IPI_1_30," _
                    & "CASE WHEN p.nbm_sh = '38151900' THEN (pvp.ValorLiquido * 0.0650) * pvp.Qtd ELSE 0 END AS IPI_6_50," _
                    & "CASE WHEN p.nbm_sh IN ('32081010', '32081020', '32089010', '32089029') THEN (pvp.ValorLiquido * 0.0375) * pvp.Qtd ELSE 0 END AS IPI_3_75," _
                    & "CASE WHEN p.nbm_sh IN('32091010','32081010', '32081020', '32089010', '32089029') THEN (pvp.ValorLiquido * 0.07) * pvp.Qtd ELSE 0 END AS ST," _
                    & "CASE WHEN p.nbm_sh NOT IN ( '27101230', '27101919', '29021990','32141020', '38151900', '32091010','32081010', '32081020', '32089010', '32089029' ) THEN 'SIM'END AS semClassificacao," _
                    & "pvp.Qtd * pvp.ValorLiquido as ValorTotal " _
                    & "FROM PreVendaNota pvn LEFT JOIN PreVendaProd pvp ON pvn.PreVendaNotaID = pvp.PreVendaNotaID LEFT JOIN produto_grade pg ON pvp.produto_gradeID = pg.produto_gradeID LEFT JOIN Produto p ON pg.produtoID = p.produtoID " _
                    & "WHERE pvn.StatusId = 13 AND pvn.dt_cancela IS NULL AND pvn.dav_numero LIKE '%" & tbPainel.Value & "'  ORDER BY p.cd_ref ASC ;"
        
        
        sqlCliente = "SELECT c.cd_cliente, c.Nome, pvn.dav_numero, c.CNPJ_CPF, c.InscricaoEstadual, c.Celular, D.cep, D.Logradouro, c.Num_End, E.Bairro, F.ds_cidade, G.ds_UF,  GE.ds_grupoecon, v.ds_nome, FORMAT(CAST(pvn.[Data] AS DATE), 'dd-MM-yyyy' ) AS Data, pvp.ds_modalidade "
        sqlCliente = sqlCliente + "FROM PreVendaNota pvn LEFT JOIN cliente c on c.clienteId = pvn.clienteId LEFT JOIN cep D ON c.cepID = D.cepID LEFT JOIN Bairro E ON D.bairroId = E.bairroId LEFT JOIN Cidade F ON E.cidadeID = F.cidadeID LEFT JOIN uf G ON G.UfID = F.UfID LEFT JOIN GrupoEcon GE ON GE.GrupoEconID = c.GrupoEconID LEFT JOIN vendedor v on pvn.vendedorID = v.vendedorID  LEFT JOIN PreVendaPgto pvp on pvp.PreVendaNotaID = pvn.PreVendaNotaID "
        sqlCliente = sqlCliente + "WHERE pvn.dav_numero LIKE '%" & tbPainel.Value & "';"
        
        sqlValTot = "SELECT SUM(pvp.Qtd * pvp.Valor_Unitario) AS valorTot FROM PreVendaNota pvn LEFT JOIN PreVendaProd pvp ON pvn.PreVendaNotaID = pvp.PreVendaNotaID WHERE pvn.StatusId = 13 AND pvn.dt_cancela IS NULL AND pvn.dav_numero LIKE '%" & tbPainel.Value & "';"
        
        
        On Error GoTo error
        
        conexaoSQLServer
            
            Set rs = getRecordset(sqlProdutos)
            
            While rs.EOF = False
                If rs.Fields("semClassificacao") = "SIM" Then
                    MsgBox "O produto " & Trim(rs.Fields("ds_completa")) & " não possui classificação fiscal configurada na planilha de orçamentos." & vbNewLine & "Por favor, procure o Administrador.", VBA.VbMsgBoxStyle.vbInformation, "Gerenciamento de Dados"
                    GoTo error
                End If
                rs.MoveNext
            Wend
            
            col = 2
            index = 0
            rs.MoveFirst
    
            Do Until rs.EOF = True
                lin = 28
                Do Until rs.EOF = True
                    Select Case index
                        Case Is = 0
                            Cells(lin, col).Value = rs.Fields("cd_ref")
                        Case Is = 1
                            Cells(lin, col).Value = Trim(rs.Fields("ds_completa"))
                            Cells(lin, col).Font.Size = 10
                            Cells(lin, col).Select
                            Range(Cells(lin, col), ActiveCell.offset(0, 4)).Merge
                        Case Is = 2
                            Cells(lin, col).Value = rs.Fields("Qtd")
                        Case Is = 3
                            Cells(lin, col).Value = rs.Fields("Valor_Unitario")
                            Cells(lin, col).Style = "Currency"
                        Case Is = 4
                            Cells(lin, col).Value = rs.Fields("ValorLiquido")
                            Cells(lin, col).Style = "Currency"
                        Case Is = 5
                            If optLucroReal.Value = True Then
                                Cells(lin, col).Value = rs.Fields("ST") + rs.Fields("IPI_1_30") + rs.Fields("IPI_6_50") + rs.Fields("IPI_3_75")
                                Cells(lin, col).Style = "Currency"
                                
                            ElseIf optSimplesNacional.Value = True Then
                                Cells(lin, col).Value = rs.Fields("ST")
                                Cells(lin, col).Style = "Currency"
                            End If
                        Case Is = 6
                            offset = col - 1
                            Cells(lin, col).Value = rs.Fields("ValorTotal") + Cells(lin, offset)
                            Cells(lin, col).Style = "Currency"
                        End Select
                        lin = lin + 1
                        rs.MoveNext
                Loop
                
                index = index + 1
                col = col + 1
                If index = 2 Then col = col + 4
                If index < 10 Then rs.MoveFirst
            Loop
            
            Set rs = getRecordset(sqlValTot)
            
            valorTot = rs.Fields(0)
            
            Set rs = getRecordset(sqlCliente)
            
            Set docVenda = Range("I2:L4")
        
            docVenda.Select
    
            ActiveCell.FormulaR1C1 = "ORÇAMENTO " & vbNewLine & "Nº DO DOCUMENTO " & Trim(rs.Fields("dav_numero"))
            tbRazaoSocial.Text = "(" & rs.Fields("cd_cliente") & ")" & Trim(rs.Fields("Nome"))
            tbCPFCNPJ.Text = Trim(rs.Fields("CNPJ_CPF"))
            tbIE.Text = Trim(rs.Fields("InscricaoEstadual"))
            tbEndereco.Text = "Endereço: " & Trim(rs.Fields("Logradouro"))
            tbNumero.Text = "N° " & Trim(rs.Fields("Num_End"))
            tbCidade.Text = "Cidade: " & Trim(rs.Fields("ds_cidade")) & " - " & Trim(rs.Fields("ds_UF"))
            tbBairro.Text = "Bairro: " & Trim(rs.Fields("Bairro"))
            tbTel.Text = "Tel: " & Trim(rs.Fields("Celular"))
            tbEmissao.Text = rs.Fields("Data")
            tbDigitador.Text = Trim(rs.Fields("ds_nome"))
            tbRepresentante.Text = Trim(rs.Fields("ds_grupoecon"))
            tbPagamento.Text = Trim(rs.Fields("ds_modalidade"))
            
        fechaBancoDeDados
        
        Call relatorioDeTotais(valorTot)
        
        If chbEstoque.Value = True Then Call verificaEstoque
        
        ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
        changeEnabled (False)
        Application.ScreenUpdating = True
        
    End If
    
    Exit Sub
    
error:
    If rs.EOF Then
        MsgBox "O Nº Dav digitado não pertence a um orçamento, tente novamente!", VBA.VbMsgBoxStyle.vbInformation, "Gerenciamento de Dados"
    Else
        MsgBox Err.Description, VBA.VbMsgBoxStyle.vbCritical, "Gerenciamento de Dados"
    End If
    fechaBancoDeDados
    limparInformacoes
    changeEnabled (False)
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True
    Application.ScreenUpdating = True
    
End Sub

Private Sub btnPDF_Click()
 Call ExportarPDF(tbRazaoSocial.Value)
End Sub

Sub ExportarPDF(ByVal nomeCliente As String)
'
' ExportarPDF
'

'
        changeEnabled (True)
        
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
        ThisWorkbook.Path & "/" & nomeCliente & ".pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        
        changeEnabled (False)
End Sub

Function limparInformacoes()
    
    Dim docVenda As Range
    
    Set docVenda = Range("I2:L4")
    
    docVenda.Select
    ActiveCell.FormulaR1C1 = "ORÇAMENTO " & vbNewLine & "Nº DO DOCUMENTO "
    
    For controle = 1 To ActiveSheet.OLEObjects.Count
         If ActiveSheet.OLEObjects(controle).PrintObject = True Then ActiveSheet.OLEObjects(controle).Object.Value = ""
    Next
    
    col = 2
    
    Do Until Cells(28, col) = ""
        lin = 28
        Do Until Cells(lin, col) = ""
            Cells(lin, col).Select
            If col = 3 Then Selection.MergeCells = False
            Selection.WrapText = False
            Rows(lin).RowHeight = 15
            Selection.Clear
            lin = lin + 1
        Loop
        If col = 3 Then col = col + 4
        col = col + 1
    Loop

End Function