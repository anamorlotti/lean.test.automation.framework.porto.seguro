'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 07/03/2018 21:11:40 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Web Browser
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Consulta de Erros e Validações
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Consulta_de_Erros_e_Validacoes
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Consulta_de_Erros_e_Validacoes
    Public Class test_KIPREV_Consulta_de_Erros_e_Validacoes
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                    	Try
                    		Test.WaitExist("27,428,94,20;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180306180738.bmp", typeIdentification.leantest) 'click Consistir
                            Test.TestLog("Acessar a funcionalidade Consulta de Erros e Validações", "Sistema deve apresentar a funcionalidade Consulta de Erros e Validações", "Sistema apresentou a funcionalidade Consulta de Erros e Validações com sucesso", typelog.Passed)
                            
                            If CBool(vbtnConsistir) Then
								'Test.TestLog("Clicar em Consistir", "Clicar em Consistir e verificar o resultado esperado", "Clique em Consistir com sucesso", typelog.Passed)
								Test.Click("27,428,94,20;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180306180738.bmp", vbtnConsistir, typeIdentification.leantest) 'click Consistir
                                'Test.WaitExist("409,428,100,22;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180306181357.bmp", typeIdentification.leantest) 'Esperar botão Efetivar
                                Test.TestLog("Resultado após clique em Consistir", "Resultado após clique em Consistir", "Resultado verificado com sucesso", typelog.Passed)
							End If
							If CBool(vbtnEfetivar) Then
								'Test.TestLog("Clicar em Efetivar", "Clicar em Efetivar e verificar o resultado esperado", "Clique em Efetivar com sucesso", typelog.Passed)
								Test.Click("409,428,100,22;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180306181357.bmp", vbtnEfetivar, typeIdentification.leantest) 'click Efetivar
								Test.TestLog("Resultado após clique em Efetivar", "Resultado após clique em Efetivar", "Resultado verificado com sucesso", typelog.Passed)
							End If
							If CBool(vbtnImprimir) Then
								'Test.TestLog("Clicar em Imprimir", "Clicar em Imprimir e verificar o resultado esperado", "Clique em Imprimir com sucesso", typelog.Passed)
								Test.Click("560,431,51,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180306181714.bmp", vbtnImprimir, typeIdentification.leantest) 'click Imprimir
								Test.TestLog("Resultado após clique em Imprimir", "Resultado após clique em Imprimir", "Resultado verificado com sucesso", typelog.Passed)
							End If
							
							'Test.Click("710,49,18,21;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180316101023.bmp", True, typeIdentification.leantest)'Fechar janela
							
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
				HandlerError("test_KIPREV_Consulta_de_Erros_e_Validacoes.test_KIPREV_Consulta_de_Erros_e_Validacoes.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vbtnConsistir,vbtnEfetivar,vbtnImprimir, vIsOpenSystem As String

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
                    vbtnConsistir = pc_db.Fieldt("vbtnConsistir")
					vbtnEfetivar = pc_db.Fieldt("vbtnEfetivar")
					vbtnImprimir = pc_db.Fieldt("vbtnImprimir")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Consulta_de_Erros_e_Validacoes.test_KIPREV_Consulta_de_Erros_e_Validacoes.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
