'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 21/05/2018 15:28:24 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to File
'Environment:..... Automation Project
'Application...... XP Parceiro
'Functionality:... Movi
'Master Test:..... No Defined
'TableTest:....... test_XP_Parceiro_Movi
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_XP_Parceiro_Movi
    Public Class test_XP_Parceiro_Movi
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                    	Try
                    	
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\Files\XP Parceiros\XPMOVI\XPMOVI230301.TXT"'xml.Read(""urlLocal"") 'Create urlLocal element in XML
								Test.TestLog("Abrir o arquivo XP Parceiro", "O arquivo  Cada deve ser aberto para leitura", "O arquivo Cada foi aberto com sucesso", typelog.Passed, , False)
							End If
                            Test.ValidadeFileField(p_pathUrlApp, "0;;1;1;Numeric", "Tipo_de_Registro_0",vTipo_de_Registro_0) 'validate filedTipo de Registro 0
							Test.ValidadeFileField(p_pathUrlApp, "0;;2;9;Date", "Data_atual",vData_atual) 'validate filedData atual
							Test.ValidadeFileField(p_pathUrlApp, "0;;10;31;Numeric", "Numero_Sequencial_do_Arquivo",vNumero_Sequencial_do_Arquivo) 'validate filedNumero Sequencial do Arquivo
							Test.ValidadeFileField(p_pathUrlApp, "6;;32;32;Numeric", "Tipo_de_Registro_6",vTipo_de_Registro_6) 'validate filedTipo de Registro 6
							Test.ValidadeFileField(p_pathUrlApp, "6;;33;36;Alphanumeric", "Historico",vHistorico) 'validate filedHistórico
							Test.ValidadeFileField(p_pathUrlApp, "6;;37;48;Numeric", "Proposta",vProposta) 'validate filedProposta
							Test.ValidadeFileField(p_pathUrlApp, "6;;49;60;Numeric", "Certificado",vCertificado) 'validate filedCertificado
							Test.ValidadeFileField(p_pathUrlApp, "6;;61;80;Alphanumeric", "Proc_SUSEP",vProc_SUSEP) 'validate filedProc SUSEP
							Test.ValidadeFileField(p_pathUrlApp, "6;;81;140;Alphanumeric", "Nome_do_Fundo",vNome_do_Fundo) 'validate filedNome do Fundo
							Test.ValidadeFileField(p_pathUrlApp, "6;;141;145;Numeric", "Codigo_do_Fundo",vCodigo_do_Fundo) 'validate filedCodigo do Fundo
							Test.ValidadeFileField(p_pathUrlApp, "6;;146;153;Numeric", "Conta_Fundo",vConta_Fundo) 'validate filedConta Fundo
							Test.ValidadeFileField(p_pathUrlApp, "6;;154;161;Date", "Data_Movimento",vData_Movimento) 'validate filedData Movimento
							Test.ValidadeFileField(p_pathUrlApp, "6;;162;179;Numeric", "Quantidade_de_Cotas",vQuantidade_de_Cotas) 'validate filedQuantidade de Cotas
							Test.ValidadeFileField(p_pathUrlApp, "6;;180;190;Numeric", "Valor_do_Movimento",vValor_do_Movimento) 'validate filedValor do Movimento
							Test.ValidadeFileField(p_pathUrlApp, "6;;191;198;Date", "Data_da_Cota",vData_da_Cota) 'validate filedData da Cota
							Test.ValidadeFileField(p_pathUrlApp, "6;;199;216;Numeric", "Valor_da_Cota",vValor_da_Cota) 'validate filedValor da Cota
							Test.ValidadeFileField(p_pathUrlApp, "6;;217;227;Numeric", "Valor_da_DA",vValor_da_DA) 'validate filedValor da DA
							Test.ValidadeFileField(p_pathUrlApp, "6;;228;238;Alphanumeric", "Valor_do_IR",vValor_do_IR) 'validate filedValor do IR
							Test.ValidadeFileField(p_pathUrlApp, "6;;239;250;Numeric", "Numero_do_Documento",vNumero_do_Documento) 'validate filedNumero do Documento
							Test.ValidadeFileField(p_pathUrlApp, "6;;251;258;Date", "Data_do_pagamento",vData_do_pagamento) 'validate filedData do pagamento
							Test.ValidadeFileField(p_pathUrlApp, "6;;259;268;Numeric", "Numero_da_parcela",vNumero_da_parcela) 'validate filedNumero da parcela
							Test.ValidadeFileField(p_pathUrlApp, "6;;269;280;Numeric", "Taxa_de_saida",vTaxa_de_saida) 'validate filedTaxa de saída
							Test.ValidadeFileField(p_pathUrlApp, "6;;281;310;Alphabetic", "Descicao_da_operacao_realizada",vDescicao_da_operacao_realizada) 'validate filedDescição da operação realizada
							Test.ValidadeFileField(p_pathUrlApp, "6;;311;324;Numeric", "Codigo_do_angaridor",vCodigo_do_angaridor) 'validate filedCodigo do angaridor
							Test.ValidadeFileField(p_pathUrlApp, "6;;325;332;Numeric", "Agencia_ou_ponto_de_venda",vAgencia_ou_ponto_de_venda) 'validate filedAgencia ou ponto de venda
							Test.ValidadeFileField(p_pathUrlApp, "6;;333;352;Numeric", "Identificacao_do_cliente_no_parceiro",vIdentificacao_do_cliente_no_parceiro) 'validate filedIdentificação do cliente no parceiro
							Test.ValidadeFileField(p_pathUrlApp, "6;;353;382;Alphabetic", "Nome_do_produto_que_foi_vendido",vNome_do_produto_que_foi_vendido) 'validate filedNome do produto que foi vendido
							Test.ValidadeFileField(p_pathUrlApp, "6;;383;393;Numeric", "Valor_nominal",vValor_nominal) 'validate filedValor nominal
							Test.ValidadeFileField(p_pathUrlApp, "6;;394;407;Numeric", "CGC_CPF_do_cliente",vCGC_CPF_do_cliente) 'validate filedCGC/CPF do cliente
							Test.ValidadeFileField(p_pathUrlApp, "6;;408;416;Numeric", "Codigo_do_movimento",vCodigo_do_movimento) 'validate filedCodigo do movimento
							Test.ValidadeFileField(p_pathUrlApp, "6;;417;424;Date", "Data_da_solicitacao",vData_da_solicitacao) 'validate filedData da solicitação
							Test.ValidadeFileField(p_pathUrlApp, "T;;425;425;Alphanumeric", "Tipo_de_Registro_T",vTipo_de_Registro_T) 'validate filedTipo de Registro T
							Test.ValidadeFileField(p_pathUrlApp, "T;;426;435;Numeric", "Quantidade_de_Linhas_do_Aquivo",vQuantidade_de_Linhas_do_Aquivo) 'validate filedQuantidade de Linhas do Aquivo
							
                            'Checkpoint
                            Test.CheckPointTest(p_CheckPoint1, p_ExpectedResult)
                            'end test                         
                            Test.EndTest(p_GenerateLogTest)
                            If p_IsLoop Then StartTest() Else p_CountTest = 0
                        Catch ex As Exception
							p_errorDescription = "Menssage error: " & ex.Message.ToString
							Test.TestLog("Passo executado", "Execução do passo com sucesso", "Passo executado com falha! Message: " & p_errorDescription, typelog.Failed)
							EndTestTable()
                       Test.EndTest(p_GenerateLogTest)
                            If p_IsLoop Then StartTest() Else p_CountTest = 0
                        End Try
                    Loop
                    EndTestTable()
                    Return True
                Else
                    Test.TestLog("Teste executado", "Teste executado com sucesso", "Teste executado com falha! StartTest = False", typelog.Failed)
                    EndTestTable()
                    Return False
                End If
            Catch ex As Exception
                p_errorDescription = "Menssage error: " & ex.Message.ToString
				HandlerError("test_XP_Parceiro_Movi.test_XP_Parceiro_Movi.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vIsOpenSystem, vTipo_de_Registro_0, vData_atual, vNumero_Sequencial_do_Arquivo, vTipo_de_Registro_6, vHistorico, vProposta, vCertificado, vProc_SUSEP, vNome_do_Fundo, vCodigo_do_Fundo, vConta_Fundo, vData_Movimento, vQuantidade_de_Cotas, vValor_do_Movimento, vData_da_Cota, vValor_da_Cota, vValor_da_DA, vValor_do_IR, vNumero_do_Documento, vData_do_pagamento, vNumero_da_parcela, vTaxa_de_saida, vDescicao_da_operacao_realizada, vCodigo_do_angaridor, vAgencia_ou_ponto_de_venda, vIdentificacao_do_cliente_no_parceiro, vNome_do_produto_que_foi_vendido, vValor_nominal, vCGC_CPF_do_cliente, vCodigo_do_movimento, vData_da_solicitacao, vTipo_de_Registro_T, vQuantidade_de_Linhas_do_Aquivo As String

        Private Function StartTest() As Boolean
            Dim strQueryOut1, strQueryOut2, strQueryOut3, strQueryOut4, strQueryOut5, strQueryOut6 as string
            Try
                p_CountTest = pc_db.OpenTestTable(p_TableTest, p_IDScenario) 'opening the test table containing all the test cases
                If p_CountTest <> 0 Then
                    p_IDScenario = pc_db.Fieldt("IDScenario") 'set IDSceario
                    p_IDTest = pc_db.Fieldt("IDTest") 'set IDTest
                    p_OrdemTest = pc_db.Fieldt("Ordem")
                    p_TestName = pc_db.Fieldt("TestName")
                    p_DescriptionTest = pc_db.Fieldt("Description")
                    p_IDRun = pc_db.Fieldt("IDRun")
                    p_ExpectedResult = pc_db.Fieldt("ExpectedResult")
                    p_IDTestInstance = pc_db.Fieldt("IDTool")
					p_CheckPoint1 = pc_db.Fieldt("CheckPoint1")

                    'Data Transfer Parameters
                    strQueryOut1 = pc_db.Fieldt("QueryInput1")
                    strQueryOut2 = pc_db.Fieldt("QueryInput2")
                    strQueryOut3 = pc_db.Fieldt("QueryInput3")
                    strQueryOut4 = pc_db.Fieldt("QueryInput4")
					strQueryOut5 = pc_db.Fieldt("QueryInput5")
					strQueryOut6 = pc_db.Fieldt("QueryInput6")
                   
                    'transfer values between tables
					If String.IsNullOrEmpty(strQueryOut1) Then pc_db.TransferDataInTablesArray(strQueryOut1, p_TableTest, p_IDScenario, p_IDTest)
                    If String.IsNullOrEmpty(strQueryOut2) Then pc_db.TransferDataInTablesArray(strQueryOut1, p_TableTest, p_IDScenario, p_IDTest)
                    If String.IsNullOrEmpty(strQueryOut3) Then pc_db.TransferDataInTablesArray(strQueryOut1, p_TableTest, p_IDScenario, p_IDTest)
                    If String.IsNullOrEmpty(strQueryOut4) Then pc_db.TransferDataInTablesArray(strQueryOut1, p_TableTest, p_IDScenario, p_IDTest)
                    If String.IsNullOrEmpty(strQueryOut5) Then pc_db.TransferDataInTablesArray(strQueryOut1, p_TableTest, p_IDScenario, p_IDTest)
                    If String.IsNullOrEmpty(strQueryOut6) Then pc_db.TransferDataInTablesArray(strQueryOut1, p_TableTest, p_IDScenario, p_IDTest)
					'
                    p_CountTest = pc_db.OpenTestTable(p_TableTest, p_IDScenario)
                    vTipo_de_Registro_0 = pc_db.Fieldt("vTipo_de_Registro_0")
					vData_atual = pc_db.Fieldt("vData_atual")
					vNumero_Sequencial_do_Arquivo = pc_db.Fieldt("vNumero_Sequencial_do_Arquivo")
					vTipo_de_Registro_6 = pc_db.Fieldt("vTipo_de_Registro_6")
					vHistorico = pc_db.Fieldt("vHistorico")
					vProposta = pc_db.Fieldt("vProposta")
					vCertificado = pc_db.Fieldt("vCertificado")
					vProc_SUSEP = pc_db.Fieldt("vProc_SUSEP")
					vNome_do_Fundo = pc_db.Fieldt("vNome_do_Fundo")
					vCodigo_do_Fundo = pc_db.Fieldt("vCodigo_do_Fundo")
					vConta_Fundo = pc_db.Fieldt("vConta_Fundo")
					vData_Movimento = pc_db.Fieldt("vData_Movimento")
					vQuantidade_de_Cotas = pc_db.Fieldt("vQuantidade_de_Cotas")
					vValor_do_Movimento = pc_db.Fieldt("vValor_do_Movimento")
					vData_da_Cota = pc_db.Fieldt("vData_da_Cota")
					vValor_da_Cota = pc_db.Fieldt("vValor_da_Cota")
					vValor_da_DA = pc_db.Fieldt("vValor_da_DA")
					vValor_do_IR = pc_db.Fieldt("vValor_do_IR")
					vNumero_do_Documento = pc_db.Fieldt("vNumero_do_Documento")
					vData_do_pagamento = pc_db.Fieldt("vData_do_pagamento")
					vNumero_da_parcela = pc_db.Fieldt("vNumero_da_parcela")
					vTaxa_de_saida = pc_db.Fieldt("vTaxa_de_saida")
					vDescicao_da_operacao_realizada = pc_db.Fieldt("vDescicao_da_operacao_realizada")
					vCodigo_do_angaridor = pc_db.Fieldt("vCodigo_do_angaridor")
					vAgencia_ou_ponto_de_venda = pc_db.Fieldt("vAgencia_ou_ponto_de_venda")
					vIdentificacao_do_cliente_no_parceiro = pc_db.Fieldt("vIdentificacao_do_cliente_no_parceiro")
					vNome_do_produto_que_foi_vendido = pc_db.Fieldt("vNome_do_produto_que_foi_vendido")
					vValor_nominal = pc_db.Fieldt("vValor_nominal")
					vCGC_CPF_do_cliente = pc_db.Fieldt("vCGC_CPF_do_cliente")
					vCodigo_do_movimento = pc_db.Fieldt("vCodigo_do_movimento")
					vData_da_solicitacao = pc_db.Fieldt("vData_da_solicitacao")
					vTipo_de_Registro_T = pc_db.Fieldt("vTipo_de_Registro_T")
					vQuantidade_de_Linhas_do_Aquivo = pc_db.Fieldt("vQuantidade_de_Linhas_do_Aquivo")
					vIsOpenSystem =pc_db.Fieldt("vIsOpenSystem")
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_XP_Parceiro_Movi.test_XP_Parceiro_Movi.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
