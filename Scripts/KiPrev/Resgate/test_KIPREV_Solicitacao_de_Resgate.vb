'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 10/03/2018 17:58:07 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Client Server (Desktop)
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Solicitação de Resgate
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Solicitacao_de_Resgate
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Solicitacao_de_Resgate
    Public Class test_KIPREV_Solicitacao_de_Resgate
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
							Test.Set("29,204,143,27;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310173505.bmp", vCertificado,"", typeIdentification.leantest) 'type Certificado
							Test.Select("483,232,164,13;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310173551.bmp", vBruto_Liq, typeIdentification.leantest) 'select Bruto Liq
							Test.SendKey("{TAB}")
							Test.Select("554,249,81,61;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310173652.bmp", vRetirada, typeIdentification.leantest) 'select Retirada
							Test.SendKey("{TAB}")
							Test.SendKey("{Delete}")		
							Test.Set("639,252,116,54;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310173755.bmp", vValor,"", typeIdentification.leantest) 'type Valor
							If CBool(vbtnCalcular_Liquido) Then
								Test.TestLog("Clicar em Calcular  Liquido", "Clicar em Calcular  Liquido e verificar o resultado esperado", "Clique em Calcular  Liquido com sucesso", typelog.Passed)
								Test.Click("238,345,137,24;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310174026.bmp", vbtnCalcular_Liquido, typeIdentification.leantest) 'click Calcular  Liquido
								Test.TestLog("Resultado após clique em Calcular  Liquido", "Resultado após clique em Calcular  Liquido", "Resultado verificado com sucesso", typelog.Passed)
							End If
							Test.WaitExist("17,400,40,56;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310174132.bmp", typeIdentification.leantest) 'type Forma de Pagamento
							Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
							Test.Set("17,400,40,56;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310174132.bmp", vForma_de_Pagamento,"", typeIdentification.leantest) 'type Forma de Pagamento
							Test.SendKey("{TAB}")
							Test.Set("222,400,47,51;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310174728.bmp", vBanco,"", typeIdentification.leantest) 'type Banco
							Test.SendKey("{TAB}")
							If CBool(vbtnValidar) Then
								Test.TestLog("Clicar em Validar", "Clicar em Validar e verificar o resultado esperado", "Clique em Validar com sucesso", typelog.Passed)
								Test.Click("136,478,78,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310175025.bmp", vbtnValidar, typeIdentification.leantest) 'click Validar
								Test.TestLog("Resultado após clique em Validar", "Resultado após clique em Validar", "Resultado verificado com sucesso", typelog.Passed)
							End If
							If CBool(vbtnOk) Then
								Test.TestLog("Clicar em Ok", "Clicar em Ok e verificar o resultado esperado", "Clique em Ok com sucesso", typelog.Passed)
								Test.Click("624,507,64,24;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310175231.bmp", vbtnOk, typeIdentification.leantest) 'click Ok
								Test.TestLog("Resultado após clique em Ok", "Resultado após clique em Ok", "Resultado verificado com sucesso", typelog.Passed)
							End If
							If CBool(vbtnErros) Then
								Test.TestLog("Clicar em Erros", "Clicar em Erros e verificar o resultado esperado", "Clique em Erros com sucesso", typelog.Passed)
								Test.Click("211,477,66,21;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310175354.bmp", vbtnErros, typeIdentification.leantest) 'click Erros
										Test.TestLog("Resultado após clique em Erros", "Resultado após clique em Erros", "Resultado verificado com sucesso", typelog.Passed)
							End If
							'If CBool(vbtnProtocolo) Then
							'	Test.TestLog("Clicar em Protocolo", "Clicar em Erros e verificar o resultado esperado", "Clique em Protocolo com sucesso", typelog.Passed)
							'	Test.Click("15,477,119,20;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310185726.bmp", vbtnProtocolo, typeIdentification.leantest) 'click Protocolo 
							'	Test.TestLog("Resultado após clique em Protocolo", "Resultado após clique em Protocolo", "Resultado verificado com sucesso", typelog.Passed)
							'End If
							
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
				HandlerError("test_KIPREV_Solicitacao_de_Resgate.test_KIPREV_Solicitacao_de_Resgate.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vCertificado, vBruto_Liq,vRetirada,vValor, vbtnCalcular_Liquido,vForma_de_Pagamento, vBanco, vbtnValidar,vbtnOk,vbtnErros As String

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
                    vCertificado = pc_db.Fieldt("vCertificado")
					vBruto_Liq = pc_db.Fieldt("vBruto_Liq")
					vRetirada = pc_db.Fieldt("vRetirada")
					vValor = pc_db.Fieldt("vValor")
					vbtnCalcular_Liquido = pc_db.Fieldt("vbtnCalcular_Liquido")
					vForma_de_Pagamento = pc_db.Fieldt("vForma_de_Pagamento")
					vBanco = pc_db.Fieldt("vBanco")
					vbtnValidar = pc_db.Fieldt("vbtnValidar")
					vbtnOk = pc_db.Fieldt("vbtnOk")
					vbtnErros = pc_db.Fieldt("vbtnErros")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Solicitacao_de_Resgate.test_KIPREV_Solicitacao_de_Resgate.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
