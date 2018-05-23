'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 13/03/2018 11:20:05 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Client Server (Desktop)
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Consulta de Propostas Pendentes
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Consulta_de_Propostas_Pendentes
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Consulta_de_Propostas_Pendentes
    Public Class test_KIPREV_Consulta_de_Propostas_Pendentes
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            'Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
							Test.Set("437,178,99,50;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313110917.bmp", vCPF_Proponente,"", typeIdentification.leantest) 'type CPF Proponente
							Test.Set("32,217,150,47;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180314085810.bmp",vNumero_Proposta_trans, "", typeIdentification.leantest) 'type Numero da Proposta
							If CBool(vbtnConsulta) Then
								'Test.TestLog("Clicar em Consulta", "Clicar em Consulta e verificar o resultado esperado", "Clique em Consulta com sucesso", typelog.Passed)
								Test.Click("676,226,79,22;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313111346.bmp", vbtnConsulta, typeIdentification.leantest) 'click Consulta
								Test.TestLog("Resultado após clique em Consulta", "Resultado após clique em Consulta", "Resultado verificado com sucesso", typelog.Passed)
							End If
							'Test.Select("88,261,43,13;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313111617.bmp", vProposta, typeIdentification.leantest) 'select Proposta
							'Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
							If CBool(vbtnNavegar) Then
								'Test.TestLog("Clicar em Navegar", "Clicar em Navegar e verificar o resultado esperado", "Clique em Navegar com sucesso", typelog.Passed)
								Test.Click("623,484,129,24;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313111729.bmp", vbtnNavegar, typeIdentification.leantest) 'click Navegar
								Test.TestLog("Resultado após clique em Navegar", "Resultado após clique em Navegar", "Resultado verificado com sucesso", typelog.Passed)
							End If
							
							Test.Click("710,49,18,21;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180316101023.bmp", True, typeIdentification.leantest)'Fechar janela
							
							
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
				HandlerError("test_KIPREV_Consulta_de_Propostas_Pendentes.test_KIPREV_Consulta_de_Propostas_Pendentes.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vNumero_Proposta_trans, vCPF_Proponente, vbtnConsulta,vProposta,vbtnNavegar As String

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
                    vCPF_Proponente = pc_db.Fieldt("vCPF_Proponente")
					vbtnConsulta = pc_db.Fieldt("vbtnConsulta")
					vProposta = pc_db.Fieldt("vProposta")
					vbtnNavegar = pc_db.Fieldt("vbtnNavegar")
					vNumero_Proposta_trans = pc_db.Fieldt("vNumero_Proposta_trans")
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Consulta_de_Propostas_Pendentes.test_KIPREV_Consulta_de_Propostas_Pendentes.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
