'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 10/03/2018 19:21:51 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Client Server (Desktop)
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Pesquisa de Pendênca de Resgate
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Pesquisa_de_Pendenca_de_Resgate
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Pesquisa_de_Pendenca_de_Resgate
    Public Class test_KIPREV_Pesquisa_de_Pendenca_de_Resgate
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vbtnTratar_Pendencias) Then
								Test.TestLog("Clicar em Tratar Pendências", "Clicar em Tratar Pendências e verificar o resultado esperado", "Clique em Tratar Pendências com sucesso", typelog.Passed)
								Test.Click("606,338,122,25;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310185252.bmp", vbtnTratar_Pendencias, typeIdentification.leantest) 'click Tratar Pendências
								Test.TestLog("Resultado após clique em Tratar Pendências", "Resultado após clique em Tratar Pendências", "Resultado verificado com sucesso", typelog.Passed)
							End If
							If CBool(vbtnSalvar) Then
								Test.TestLog("Clicar em Salvar", "Clicar em Salvar e verificar o resultado esperado", "Clique em Salvar com sucesso", typelog.Passed)
								Test.Click("44,433,99,17;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310191933.bmp", vbtnSalvar, typeIdentification.leantest) 'click Salvar
								Test.TestLog("Resultado após clique em Salvar", "Resultado após clique em Salvar", "Resultado verificado com sucesso", typelog.Passed)
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
				HandlerError("test_KIPREV_Pesquisa_de_Pendenca_de_Resgate.test_KIPREV_Pesquisa_de_Pendenca_de_Resgate.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vbtnTratar_Pendencias,vbtnSalvar As String

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
                    vbtnTratar_Pendencias = pc_db.Fieldt("vbtnTratar_Pendencias")
					vbtnSalvar = pc_db.Fieldt("vbtnSalvar")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Pesquisa_de_Pendenca_de_Resgate.test_KIPREV_Pesquisa_de_Pendenca_de_Resgate.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
