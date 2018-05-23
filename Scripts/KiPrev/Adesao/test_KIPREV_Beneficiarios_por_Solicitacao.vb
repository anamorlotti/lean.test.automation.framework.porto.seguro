'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 28/02/2018 17:42:11 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Client Server (Desktop)
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Beneficiarios por Solicitaçao
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Beneficiarios_por_Solicitacao
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Beneficiarios_por_Solicitacao
    Public Class test_KIPREV_Beneficiarios_por_Solicitacao
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                    	Try
                    		Test.waitexist("188,79,52,138;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180314113600.bmp", typeidentification.leanTest)
                            Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
                            Dim numero_proposta As String 
                            Do
								Test.DoubleClick("188,79,52,138;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180314113600.bmp", True, typeidentification.leanTest)
								Test.Wait(4000)
								Test.ClipBoardClear
	                            Test.SendKey("^c")
	                            numero_proposta = Test.ClipBoardGetText
	                            Test.SetValueOutput("vnumero_proposta_out", numero_proposta)
                            Loop While String.IsNullOrEmpty(numero_proposta) 
                                                      
                            Test.Set("16,213,72,63;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180301171305.bmp", vCod_Beneficiario,"", typeIdentification.leantest) 'type Cod Beneficiário
                            Test.SendKey("{TAB}")
                            Test.Set("556,216,39,55;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180301171407.bmp", vCod_Parentesco,"", typeIdentification.leantest) 'type Cod Parentesco
                            Test.SendKey("{TAB}")
                            If CBool(vbtnSalvar) Then
								Test.TestLog("Clicar em Salvar", "Clicar em Salvar e verificar o resultado esperado", "Clique em Salvar com sucesso", typelog.Passed)
								Test.Click("748,491,30,30;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180301171536.bmp", vbtnSalvar, typeIdentification.leantest) 'click Salvar
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
				HandlerError("test_KIPREV_Beneficiarios_por_Solicitacao.test_KIPREV_Beneficiarios_por_Solicitacao.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vCod_Beneficiario, vCod_Parentesco, vbtnSalvar As String

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
                    vCod_Beneficiario = pc_db.Fieldt("vCod_Beneficiario")
                    vCod_Parentesco = pc_db.Fieldt("vCod_Parentesco")
					vbtnSalvar = pc_db.Fieldt("vbtnSalvar")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Beneficiarios_por_Solicitacao.test_KIPREV_Beneficiarios_por_Solicitacao.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
