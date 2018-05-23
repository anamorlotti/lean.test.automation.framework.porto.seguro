'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 28/02/2018 12:17:38 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Web Browser
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Adesao Dados de Comercializaçao
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Dados_de_Comercializacao
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Dados_de_Comercializacao
    Public Class test_KIPREV_Dados_de_Comercializacao
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                	Do While p_CountTest <> 0
                		Try
                			''*******************************************************************************************************************************************************
							'Canal de Vendas
							Test.WaitExist("419,109,88,60;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228110034.bmp", typeidentification.leanTest)
							Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
							'Test.Set("34,115,117,48;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180306084454.bmp", vCod_Canal, "", typeidentification.leanTest)'type Cod Canal
							Test.Set("419,109,88,60;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228110034.bmp", vAgencia_Captadora,"", typeIdentification.leantest) 'type Agencia Captadora
							''*******************************************************************************************************************************************************
							'Origem de Vendas
							Test.Set("36,166,110,46;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228142125.bmp", vCod_Origem_Vendas,"", typeIdentification.leantest) 'type Cod Origem Vendas
							Test.Set("410,155,71,72;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228110145.bmp", vMatricula,"", typeIdentification.leantest) 'type Matricula
							''*******************************************************************************************************************************************************
							'Corretores
							Test.Set("39,209,81,68;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228110455.bmp", vCod_Corretor1,"", typeIdentification.leantest) 'type Cod Corretor1
							Test.Click("717,219,28,38;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228111828.bmp", vPrincipal_Corretor1, typeIdentification.leantest) 'checked Principal Corretor1
							Test.Set("39,209,81,68;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228110455.bmp", vCod_Corretor2,"", typeIdentification.leantest) 'type Cod Corretor2
							Test.Set("643,250,98,25;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228121150.bmp", vValor_Corretor2,"", typeIdentification.leantest) 'type Valor Corretor2
							Test.Click("707,250,49,28;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228121240.bmp", vPrincipal_Corretor2, typeIdentification.leantest) 'checked Principal Corretor2
							''*******************************************************************************************************************************************************
							'Final
							Test.Set("38,292,106,48;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228112241.bmp", vAgenciador,"", typeIdentification.leantest) 'type Agenciador
							Test.WaitExist("36,406,110,53;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228112551.bmp", typeIdentification.leantest) 'type Regra Corretagem
							Test.Set("36,406,110,53;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228112551.bmp", vRegra_Corretagem,"", typeIdentification.leantest) 'type Regra Corretagem
							Test.SendKey("{Tab}")	
							Test.Wait(200)
							
							If CBool(vbtnSalvar) Then
								Test.WaitExist("750,491,28,26;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228121353.bmp", typeIdentification.leantest) 'Wait exist Salvar
								Test.TestLog("Clicar em Salvar", "Clicar em Salvar e verificar o resultado esperado", "Clique em Salvar com sucesso", typelog.Passed)
								Test.Click("750,491,28,26;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228121353.bmp", vbtnSalvar, typeIdentification.leantest) 'click Salvar
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
				HandlerError("test_KIPREV_Adesao_Dados_de_Comercializacao.test_KIPREV_Adesao_Dados_de_Comercializacao.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vCod_Canal, vAgencia_Captadora, vCod_Origem_Vendas, vMatricula, vCod_Corretor1, vValor_Corretor_1, vPrincipal_Corretor1,vCod_Corretor2, vValor_Corretor2, vPrincipal_Corretor2,vAgenciador, vRegra_Corretagem, vbtnSalvar, vIsOpenSystem As String

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
					vCod_Canal = pc_db.Fieldt("vCod_Canal")
                    vAgencia_Captadora = pc_db.Fieldt("vAgencia_Captadora")
					vCod_Origem_Vendas = pc_db.Fieldt("vCod_Origem_Vendas")
					vMatricula = pc_db.Fieldt("vMatricula")
					vCod_Corretor1 = pc_db.Fieldt("vCod_Corretor1")
					vValor_Corretor_1 = pc_db.Fieldt("vValor_Corretor_1")
					vPrincipal_Corretor1 = pc_db.Fieldt("vPrincipal_Corretor1")
					vCod_Corretor2 = pc_db.Fieldt("vCod_Corretor2")
					vValor_Corretor2 = pc_db.Fieldt("vValor_Corretor2")
					vPrincipal_Corretor2 = pc_db.Fieldt("vPrincipal_Corretor2")
					vAgenciador = pc_db.Fieldt("vAgenciador")
					vRegra_Corretagem = pc_db.Fieldt("vRegra_Corretagem")
					vbtnSalvar = pc_db.Fieldt("vbtnSalvar")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Adesao_Dados_de_Comercializacao.test_KIPREV_Adesao_Dados_de_Comercializacao.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
