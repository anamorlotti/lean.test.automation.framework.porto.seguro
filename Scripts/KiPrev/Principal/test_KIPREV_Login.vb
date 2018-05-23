'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 26/02/2018 15:27:22 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Client Server (Desktop)
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Login
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Login
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Login
    Public Class test_KIPREV_Login
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								'p_pathUrlApp = "c:\users\ssuba\desktop\kiprev.lnk"'xml.Read(""urlLocal"") 'Create urlLocal element in XML
								pc_util.StopProcess("javaw")
								'Test.OpenApp(p_pathUrlApp)
								Test.StartTests
								Test.Open("http://23.22.189.78:7777/forms/frmservlet?config=pseg02")								
								'Test.OpenApp("c:\kiprev\KiPREV - PSEG02.lnk")
                            End If
                            'mensagem de bloqueio Kimam
                            'Test.WaitExist("811,399,111,30;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180226163827.bmp", typeidentification.leanTest, , 20)
                            'test.Click("811,399,111,30;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180226163827.bmp", True, typeidentification.leanTest)	
                            'fim da mensagem de bloqueio
                            
							Test.WaitExist("287,278,57,27;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180226153904.bmp", typeIdentification.leantest) 'WaitExist 
							If CBool(vbtnlogin) Then
								Test.Click("387,308,59,28;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180226152045.bmp", vbtnlogin, typeIdentification.leantest) 'click login
							End If
							Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
							Test.Set("365,306,108,31;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180226152305.bmp", vlogin1,"", typeIdentification.leantest) 'type login1
							Test.SendKey("{TAB}") 'SendKey {TAB}
							Test.SendKey(vSenha) 'SendKey Senha
							Test.SendKey("{TAB}") 'SendKey {TAB}
							Test.SendKey(vbanco_de_dados) 'SendKey banco de dados
							Test.SendKey("{TAB}") 'SendKey {TAB}
							If CBool(vbtnConectar) Then
								Test.TestLog("Clicar em Conectar", "Clicar em Conectar e verificar o resultado esperado", "Clique em Conectar com sucesso", typelog.Passed)
								Test.Click("380,397,68,29;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180226152539.bmp", vbtnConectar, typeIdentification.leantest) 'click Conectar
								Test.WaitExist("1,22,84,24;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180226170240.bmp", typeIdentification.leantest) 'click RADM KIPREV
								Test.TestLog("Resultado após clique em Conectar", "Resultado após clique em Conectar", "Resultado verificado com sucesso", typelog.Passed)
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
				HandlerError("test_KIPREV_Login.test_KIPREV_Login.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vbtnlogin,vlogin1, vSenha, vbanco_de_dados, vbtnConectar, vIsOpenSystem As String

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
                    vbtnlogin = pc_db.Fieldt("vbtnlogin")
					vlogin1 = pc_db.Fieldt("vlogin1")
					vSenha = pc_db.Fieldt("vSenha")
					vbanco_de_dados = pc_db.Fieldt("vbanco_de_dados")
					vbtnConectar = pc_db.Fieldt("vbtnConectar")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Login.test_KIPREV_Login.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
