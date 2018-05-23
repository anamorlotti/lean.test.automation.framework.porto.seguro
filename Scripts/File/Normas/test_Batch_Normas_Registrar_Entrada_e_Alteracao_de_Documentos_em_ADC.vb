'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 17/05/2018 17:57:50 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to File
'Environment:..... Automation Project
'Application...... Batch Normas
'Functionality:... Registrar Entrada e Alteração de Documentos em ADC
'Master Test:..... No Defined
'TableTest:....... test_Batch_Normas_Registrar_Entrada_e_Alteracao_de_Documentos_em
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Batch_Normas_Registrar_Entrada_e_Alteracao_de_Documentos_em
    Public Class test_Batch_Normas_Registrar_Entrada_e_Alteracao_de_Documentos_em
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\Files\Batch Normas"'xml.Read(""urlLocal"") 'Create urlLocal element in XML
								Test.TestLog("Abrir o arquivo Batch Normas", "O arquivo  Registrar Entrada e Alteração de Documentos em ADC deve ser aberto para leitura", "O arquivo Registrar Entrada e Alteração de Documentos em ADC foi aberto com sucesso", typelog.Passed, , False)
							End If
							Test.ValidadeFileField(p_pathUrlApp, ";;1;9;Alphanumeric", "Numero_da_Proposta",vNumero_da_Proposta) 'validate filedNumero da Proposta
							Test.ValidadeFileField(p_pathUrlApp, ";;10;69;Alphanumeric", "Numero_do_Contrato",vNumero_do_Contrato) 'validate filedNumero do Contrato
							Test.ValidadeFileField(p_pathUrlApp, ";;70;79;Alphanumeric", "Numero_do_Endosso",vNumero_do_Endosso) 'validate filedNumero do Endosso
							Test.ValidadeFileField(p_pathUrlApp, ";;80;82;Alphanumeric", "Numero_da_Parcela",vNumero_da_Parcela) 'validate filedNumero da Parcela
							Test.ValidadeFileField(p_pathUrlApp, ";;83;86;Alphanumeric", "Codigo_da_Empesa",vCodigo_da_Empesa) 'validate filedCodigo da Empesa
							Test.ValidadeFileField(p_pathUrlApp, ";;87;90;Alphanumeric", "Codigo_do_Ramo",vCodigo_do_Ramo) 'validate filedCodigo do Ramo
							Test.ValidadeFileField(p_pathUrlApp, ";;91;94;", "Codigo_da_Modalidade",vCodigo_da_Modalidade) 'validate filedCodigo da Modalidade
							Test.ValidadeFileField(p_pathUrlApp, ";;95;100;Alphanumeric", "Codigo_da_Susep",vCodigo_da_Susep) 'validate filedCodigo da Susep
							Test.ValidadeFileField(p_pathUrlApp, ";;101;103;Alphanumeric", "Codigo_do_Banco",vCodigo_do_Banco) 'validate filedCodigo do Banco
							Test.ValidadeFileField(p_pathUrlApp, ";;126;129;Alphanumeric", "Codigo_da_Agencia",vCodigo_da_Agencia) 'validate filedCodigo da Agencia
							Test.ValidadeFileField(p_pathUrlApp, ";;130;147;Alphanumeric", "Codigo_da_Conta_Corrente",vCodigo_da_Conta_Corrente) 'validate filedCodigo da Conta Corrente
							Test.ValidadeFileField(p_pathUrlApp, ";;148;167;Alphanumeric", "CPF_CNPJ",vCPF_CNPJ) 'validate filedCPF/CNPJ
							Test.ValidadeFileField(p_pathUrlApp, ";;181;193;Others", "Valor_do_Meio_de_Recebimento",vValor_do_Meio_de_Recebimento) 'validate filedValor do Meio de Recebimento
							Test.ValidadeFileField(p_pathUrlApp, ";;194;194;Alphanumeric", "Meio_de_Recebimento",vMeio_de_Recebimento) 'validate filedMeio de Recebimento
							
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
				HandlerError("test_Batch_Normas_Registrar_Entrada_e_Alteracao_de_Documentos_em.test_Batch_Normas_Registrar_Entrada_e_Alteracao_de_Documentos_em.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vNumero_da_Proposta, vNumero_do_Contrato, vNumero_do_Endosso, vNumero_da_Parcela, vCodigo_da_Empesa, vCodigo_do_Ramo, vCodigo_da_Modalidade, vCodigo_da_Susep, vCodigo_do_Banco, vCodigo_da_Agencia, vCodigo_da_Conta_Corrente, vCPF_CNPJ, vValor_do_Meio_de_Recebimento, vMeio_de_Recebimento, vIsOpenSystem As String

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
                    vNumero_da_Proposta = pc_db.Fieldt("vNumero_da_Proposta")
					vNumero_do_Contrato = pc_db.Fieldt("vNumero_do_Contrato")
					vNumero_do_Endosso = pc_db.Fieldt("vNumero_do_Endosso")
					vNumero_da_Parcela = pc_db.Fieldt("vNumero_da_Parcela")
					vCodigo_da_Empesa = pc_db.Fieldt("vCodigo_da_Empesa")
					vCodigo_do_Ramo = pc_db.Fieldt("vCodigo_do_Ramo")
					vCodigo_da_Modalidade = pc_db.Fieldt("vCodigo_da_Modalidade")
					vCodigo_da_Susep = pc_db.Fieldt("vCodigo_da_Susep")
					vCodigo_do_Banco = pc_db.Fieldt("vCodigo_do_Banco")
					vCodigo_da_Agencia = pc_db.Fieldt("vCodigo_da_Agencia")
					vCodigo_da_Conta_Corrente = pc_db.Fieldt("vCodigo_da_Conta_Corrente")
					vCPF_CNPJ = pc_db.Fieldt("vCPF_CNPJ")
					vValor_do_Meio_de_Recebimento = pc_db.Fieldt("vValor_do_Meio_de_Recebimento")
					vMeio_de_Recebimento = pc_db.Fieldt("vMeio_de_Recebimento")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Batch_Normas_Registrar_Entrada_e_Alteracao_de_Documentos_em.test_Batch_Normas_Registrar_Entrada_e_Alteracao_de_Documentos_em.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
