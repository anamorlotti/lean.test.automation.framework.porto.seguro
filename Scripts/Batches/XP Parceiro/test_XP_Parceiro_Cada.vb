'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 21/05/2018 12:19:47 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to File
'Environment:..... Automation Project
'Application...... XP Parceiro
'Functionality:... Cada
'Master Test:..... No Defined
'TableTest:....... test_XP_Parceiro_Cada
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_XP_Parceiro_Cada
    Public Class test_XP_Parceiro_Cada
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try 
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\Files\XP Parceiros\XPCADA\XPCADA0123032018.TXT"'xml.Read(""urlLocal"") 'Create urlLocal element in XML
								Test.TestLog("Abrir o arquivo XP Parceiro", "O arquivo  Cada deve ser aberto para leitura", "O arquivo Cada foi aberto com sucesso", typelog.Passed, , False)
							End If
							Test.ValidadeFileField(p_pathUrlApp, "0;;1;1;Numeric", "Identificacao_da_Linha_0",vIdentificacao_da_Linha_0) 'validate filedIdentificacao da Linha 0
							Test.ValidadeFileField(p_pathUrlApp, "0;;2;9;Date", "Data_atual",vData_atual) 'validate filedData atual
							Test.ValidadeFileField(p_pathUrlApp, "0;;10;29;Alphabetic", "Campo_Fixo_POSICAO",vCampo_Fixo_POSICAO) 'validate filedCampo Fixo POSICAO
							Test.ValidadeFileField(p_pathUrlApp, "0;;30;49;Alphabetic", "Campo_Fixo_Retorno",vCampo_Fixo_Retorno) 'validate filedCampo Fixo Retorno
							Test.ValidadeFileField(p_pathUrlApp, "0;;50;53;Numeric", "Parceria",vParceria) 'validate filedParceria
							Test.ValidadeFileField(p_pathUrlApp, "0;;54;75;Alphanumeric", "wnomearquivo",vwnomearquivo) 'validate filedwnomearquivo
							Test.ValidadeFileField(p_pathUrlApp, "1;;76;76;Numeric", "Identificacao_da_Linha1",vIdentificacao_da_Linha1) 'validate filedIdentificacao da Linha1
							Test.ValidadeFileField(p_pathUrlApp, "1;;77;88;Numeric", "Proposta",vProposta) 'validate filedProposta
							Test.ValidadeFileField(p_pathUrlApp, "1;;89;100;Numeric", "Contrato",vContrato) 'validate filedContrato
							Test.ValidadeFileField(p_pathUrlApp, "1;;101;112;Numeric", "Certificado",vCertificado) 'validate filedCertificado
							Test.ValidadeFileField(p_pathUrlApp, "1;;113;171;Alphabetic", "Situacao",vSituacao) 'validate filedSituação
							Test.ValidadeFileField(p_pathUrlApp, "2;;172;172;Numeric", "Identificacao_da_Linha_2",vIdentificacao_da_Linha_2) 'validate filedIdentificacao da Linha 2
							Test.ValidadeFileField(p_pathUrlApp, "20;;173;192;Numeric", "Proc_SUSEP",vProc_SUSEP) 'validate filedProc SUSEP
							Test.ValidadeFileField(p_pathUrlApp, "2;;193;252;Numeric", "Nome_do_Fundo",vNome_do_Fundo) 'validate filedNome do Fundo
							Test.ValidadeFileField(p_pathUrlApp, "2;;253;257;Numeric", "Codigo_do_Fundo",vCodigo_do_Fundo) 'validate filedCodigo do Fundo
							
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
				HandlerError("test_XP_Parceiro_Cada.test_XP_Parceiro_Cada.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vIdentificacao_da_Linha_0, vData_atual, vCampo_Fixo_POSICAO, vCampo_Fixo_Retorno, vParceria, vwnomearquivo, vIdentificacao_da_Linha1, vProposta, vContrato, vCertificado, vSituacao, vIdentificacao_da_Linha_2, vProc_SUSEP, vNome_do_Fundo, vCodigo_do_Fundo, vIsOpenSystem As String

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
                    vIdentificacao_da_Linha_0 = pc_db.Fieldt("vIdentificacao_da_Linha_0")
					vData_atual = pc_db.Fieldt("vData_atual")
					vCampo_Fixo_POSICAO = pc_db.Fieldt("vCampo_Fixo_POSICAO")
					vCampo_Fixo_Retorno = pc_db.Fieldt("vCampo_Fixo_Retorno")
					vParceria = pc_db.Fieldt("vParceria")
					vwnomearquivo = pc_db.Fieldt("vwnomearquivo")
					vIdentificacao_da_Linha1 = pc_db.Fieldt("vIdentificacao_da_Linha1")
					vProposta = pc_db.Fieldt("vProposta")
					vContrato = pc_db.Fieldt("vContrato")
					vCertificado = pc_db.Fieldt("vCertificado")
					vSituacao = pc_db.Fieldt("vSituacao")
					vIdentificacao_da_Linha_2 = pc_db.Fieldt("vIdentificacao_da_Linha_2")
					vProc_SUSEP = pc_db.Fieldt("vProc_SUSEP")
					vNome_do_Fundo = pc_db.Fieldt("vNome_do_Fundo")
					vCodigo_do_Fundo = pc_db.Fieldt("vCodigo_do_Fundo")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_XP_Parceiro_Cada.test_XP_Parceiro_Cada.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
