'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 28/02/2018 15:32:16 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Client Server (Desktop)
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Contrataçao de Beneficios
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Contratacao_de_Beneficios
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Contratacao_de_Beneficios
    Public Class test_KIPREV_Contratacao_de_Beneficios
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                    	Try
                    		
                    		Test.WaitExist("39,105,110,72;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180301162713.bmp", typeidentification.leanTest)         		  
					        Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
							Test.Set("52,197,55,43;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228152400.bmp", vBeneficios_de_Sobrevivencia,"", typeIdentification.leantest) 'type Beneficios de Sobrevivencia
							Test.Click("475,298,54,36;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228152656.bmp", v13_Renda, typeIdentification.leantest) 'checked 13 Renda
							Test.WaitExist("517,290,116,51;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228152733.bmp", typeIdentification.leantest) 'type Idade Aposentaoria
							Test.Wait(200)
							Test.Set("518,297,115,54;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180301163419.bmp", vIdade_Aposentaoria,"", typeIdentification.leantest) 'type Idade Aposentaoria
							Test.SendKey("{TAB}")
							Test.Set("634,291,119,50;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228152806.bmp", vData_Aposentadoria,"", typeIdentification.leantest) 'type Data Aposentadoria
							If CBool(vbtnSelecao_de_Riscos) Then
								'Test.TestLog("Clicar em Seleção de Riscos", "Clicar em Seleção de Riscos e verificar o resultado esperado", "Clique em Seleção de Riscos com sucesso", typelog.Passed)
								Test.DoubleClick("382,361,24,15;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180301164206.bmp", vbtnSelecao_de_Riscos, typeIdentification.leantest) 'DoubleClick Seleção de Riscos
								'Test.TestLog("Resultado após clique em Seleção de Riscos", "Resultado após clique em Seleção de Riscos", "Verificação realizada com sucesso", typelog.Passed)
							End If
							'Clicar em questionário
							Test.WaitExist("69,421,69,11;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180423114113.bmp", typeIdentification.leantest) 'esperar o questionário
							
							If CBool(vbtnQuestionario) Then
								Test.TestLog("Clicar em Questionário", "Clicar em Questionário e verificar o resultado esperado", "Clique em Questionário com sucesso", typelog.Passed)
								Test.Click("69,421,69,11;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180423114113.bmp", vbtnQuestionario, typeIdentification.leantest) 'click Questionário
								'Test.TestLog("Resultado após clique em Salvar", "Resultado após clique em Salvar", "Resultado verificado com sucesso", typelog.Passed)
								
							'Responder o questionário
							Test.WaitExist("26,252,141,36;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180423114450.bmp", typeIdentification.leantest) 'esperar o questionário
							Test.Set("25,408,53,44;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180423115004.bmp", vCod_Resposta,"", typeIdentification.leantest) 'type Cod_Resposta
							Test.SendKey("{TAB}")	
							
							'Salvar o questionário
								Test.TestLog("Clicar em Salvar", "Clicar em Salvar e verificar o resultado esperado", "Clique em Salvar com sucesso", typelog.Passed)
								Test.Click("17,49,17,21;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180425164642.bmp", True, typeIdentification.leantest) 'click Salvar
								Test.Click("710,50,22,17;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180425164724.bmp", True, typeIdentification.leantest) ' clicar em sair
								Test.WaitExist("710,50,22,17;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180425164724.bmp", typeIdentification.leantest) 'esperar pelo yes
								Test.SendKey("{ENTER}")	
							End If
							
							
							If CBool(vbtnSalvar) Then
								Test.TestLog("Clicar em Salvar", "Clicar em Salvar e verificar o resultado esperado", "Clique em Salvar com sucesso", typelog.Passed)
								Test.Click("748,489,28,27;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180301164115.bmp", vbtnSalvar, typeIdentification.leantest) 'click Salvar
								'Test.TestLog("Resultado após clique em Salvar", "Resultado após clique em Salvar", "Resultado verificado com sucesso", typelog.Passed)
							End If
							
							If CBool(vbtnSalvar) Then
								'Test.TestLog("Clicar em Salvar", "Clicar em Salvar e verificar o resultado esperado", "Clique em Salvar com sucesso", typelog.Passed)
								Test.Click("748,489,28,27;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180301164115.bmp", vbtnSalvar, typeIdentification.leantest) 'click Salvar
								'Test.TestLog("Resultado após clique em Salvar", "Resultado após clique em Salvar", "Resultado verificado com sucesso", typelog.Passed)
							End If
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
				HandlerError("test_KIPREV_Contratacao_de_Beneficios.test_KIPREV_Contratacao_de_Beneficios.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vbtnQuestionario, vCod_Resposta, vBeneficios_de_Sobrevivencia, v13_Renda,vIdade_Aposentaoria, vData_Aposentadoria, vbtnSelecao_de_Riscos,vbtnSalvar As String

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
                    vBeneficios_de_Sobrevivencia = pc_db.Fieldt("vBeneficios_de_Sobrevivencia")
					v13_Renda = pc_db.Fieldt("v13_Renda")
					vIdade_Aposentaoria = pc_db.Fieldt("vIdade_Aposentaoria")
					vData_Aposentadoria = pc_db.Fieldt("vData_Aposentadoria")
					vbtnSelecao_de_Riscos = pc_db.Fieldt("vbtnSelecao_de_Riscos")
					vbtnSalvar = pc_db.Fieldt("vbtnSalvar")
					vbtnQuestionario = pc_db.FieldT ("vbtnQuestionario")
					vCod_Resposta = pc_db.FieldT ("vCod_Resposta")
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Contratacao_de_Beneficios.test_KIPREV_Contratacao_de_Beneficios.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
