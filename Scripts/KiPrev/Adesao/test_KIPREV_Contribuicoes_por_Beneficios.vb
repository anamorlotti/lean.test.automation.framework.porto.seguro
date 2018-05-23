'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 01/03/2018 19:00:34 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Web Browser
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Contribuicoes por Beneficios
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Contribuicoes_por_Beneficios
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Contribuicoes_por_Beneficios
    Public Class test_KIPREV_Contribuicoes_por_Beneficios
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                    	Try
                    		Test.WaitExist("465,186,76,58;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180301185722.bmp", typeIdentification.leantest) 'WaitEXist Planos
							Test.TestLog("Acessar a funcionalidade Contribuicoes por Beneficios", "Sistema deve apresentar a funcionalidade Contribuicoes por Beneficios", "Sistema apresentou a funcionalidade Contribuicoes por Beneficios com sucesso", typelog.Passed)
							'If CBool  Then
							'	Test.TestLog("Clicar em Planos", "Clicar em Planos e verificar o resultado esperado", "Clique em Planos com sucesso", typelog.Passed)
							'	Test.Click("465,186,76,58;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180301185722.bmp", typeIdentification.leantest) 'click Planos
							'	Test.TestLog("Resultado após clique em Planos", "Resultado após clique em Planos", "Resultado verificado com sucesso", typelog.Passed)	
							'End If
							
							'Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
							'Test.Set("669,333,85,62;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180301185815.bmp", , typeIdentification.leantest) 'type Valor do Pagamento
														
							For i= 0 To CInt(vPlanos) - 1
								'seleciona a primeira linha
								If i < 1 Then 
									Test.Click("534,188,52,56;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180302141858.bmp", True, typeidentification.leanTest)
								ElseIf i = 1 Then'seleciona as demais linhas
									Test.Click("534,188,52,56;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180302141858.bmp", True, typeidentification.leanTest)
									Test.SendKey("{DOWN}", i)
								Elseif i > 1  Then'seleciona as demais linhas
									Test.Click("530,186,59,59;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180302143356.bmp", True, typeidentification.leanTest)
									Test.SendKey("{DOWN}", i)	
								End If
														
								'If i > 0 Then
								'	Test.SendKey("{DOWN}", i)
								'End If
								
								'beneficios
								Dim valorPago As String() = vValorPago.Split(";")
								If i < 1 Then
									Test.Click("26,337,65,47;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180518143951.bmp", True, typeidentification.leanTest)
									Test.DoubleClick("26,337,65,47;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180518143951.bmp", True, typeidentification.leanTest)
									Test.WaitExist("6,118,223,24;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180306142928.bmp", typeidentification.leanTest)
									Test.Set("6,118,223,24;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180306142928.bmp", vTipoContribuicao, "", Typeidentification.leanTest)
									Test.Sendkey("{TAB}")	
									Test.Click("92,442,77,28;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180302143003.bmp", True, Typeidentification.leanTest)							
									Test.Set("677,326,81,70;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180522174751.bmp", valorPago(i), "", typeidentification.leanTest)
																	
								ElseIf i = 1 Then
									Test.WaitExist("657,191,105,82;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180523104850.bmp", typeidentification.leanTest)
									Test.click("657,191,105,82;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180523104850.bmp", True, typeidentification.leanTest) 'Escrever o valor do beneficio para morte
									Test.set("657,191,105,82;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180523104850.bmp", vValorBeneficio, "", typeidentification.leanTest) 'Escrever o valor do beneficio para morte
									'Test.Sendkey("{vValorBeneficio}") 'Escrever o valor do beneficio para morte
									'Test.Set("669,334,88,62;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180302143135.bmp", valorPago(i), "", typeidentification.leanTest)
								Elseif i > 1 Then
									Test.Set("674,331,84,61;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180518152631.bmp", valorPago(i), "", typeidentification.leanTest)
									End If
								Test.TestLog("Acessar preencher Contribuicoes", "Sistema deve permitir o preenchimento Contribuicoes", "Sistema apresentou a funcionalidade Contribuicoes com sucesso", typelog.Passed)
							Next
		
							If CBool(vbtnSalvar) Then
								Test.TestLog("Clicar em Salvar", "Clicar em Salvar e verificar o resultado esperado", "Clique em Salvar com sucesso", typelog.Passed)
								Test.Click("750,497,26,20;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180301185851.bmp", vbtnSalvar, typeIdentification.leantest) 'click Salvar
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
				HandlerError("test_KIPREV_Contribuicoes_por_Beneficios.test_KIPREV_Contribuicoes_por_Beneficios.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vValorBeneficio, vPlanos, vTipoContribuicao, vValorPago, vbtnSalvar As String

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
                  	vValorBeneficio = pc_db.Fieldt("vValorBeneficio")
                    vPlanos = pc_db.Fieldt("vPlanos")
					vTipoContribuicao = pc_db.Fieldt("vTipoContribuicao")
					vValorPago = pc_db.Fieldt("vValorPago")
					vbtnSalvar = pc_db.Fieldt("vbtnSalvar")
					
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Contribuicoes_por_Beneficios.test_KIPREV_Contribuicoes_por_Beneficios.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
