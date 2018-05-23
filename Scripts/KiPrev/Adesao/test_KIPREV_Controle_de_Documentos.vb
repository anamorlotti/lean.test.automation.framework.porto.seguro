'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 13/03/2018 09:42:13 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Client Server (Desktop)
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Controle de Documentos
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Controle_de_Documentos
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Controle_de_Documentos
    Public Class test_KIPREV_Controle_de_Documentos
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                    	Try
                    		Test.WaitExist("65,128,116,20;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313093305.bmp", typeIdentification.leantest) 'type Numero Documento
							Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
							Test.Set("65,128,116,20;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313093305.bmp", vFilial,"", typeIdentification.leantest) 'waitExist Filial
							Test.SendKey("{TAB}")
							Test.WaitExist("417,239,209,25;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313093644.bmp", typeIdentification.leantest) 'type Numero Documento
							Test.Set("417,239,209,25;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313093644.bmp", vNumero_Documento,"", typeIdentification.leantest) 'type Numero Documento
							Test.Set("417,239,209,25;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313093644.bmp", vNumero_Proposta_trans,"", typeIdentification.leantest) 'type Numero Documento
							If CBool(vbtnConsultar) Then
								'Test.TestLog("Clicar em Consultar", "Clicar em Consultar e verificar o resultado esperado", "Clique em Consultar com sucesso", typelog.Passed)
								Test.Click("656,383,116,20;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313124030.bmp", vbtnConsultar, typeIdentification.leantest) 'click Consultar
								Test.TestLog("Resultado após clique em Consultar", "Resultado após clique em Consultar", "Resultado verificado com sucesso", typelog.Passed)
							End If
							If CBool(vbtnChecked_Documentos) Then
								'Test.TestLog("Clicar em Checked Documenotos", "Clicar em Checked Documentos e verificar o resultado esperado", "Clique em Checked Documentos com sucesso", typelog.Passed)
								Test.Click("720,405,33,23;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180314083711.bmp", vbtnChecked_Documentos, typeIdentification.leantest) 'click Checked Documentos
								'Test.Click("725,321,21,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313165233.bmp", vbtnChecked_Documentos, typeIdentification.leantest) 'click Checked Documentos
							End If
							If CBool(vbtnBaixar) Then
								Test.Click("493,516,31,14;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313135518.bmp", vbtnBaixar, typeIdentification.leantest) 'click Baixar
								Test.TestLog("Clicar em Baixar", "Clicar em Baixar e verificar o resultado esperado", "Clique em Baixar com sucesso", typelog.Passed)
                            End If
							'sair da tela
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
				HandlerError("test_KIPREV_Controle_de_Documentos.test_KIPREV_Controle_de_Documentos.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vbtnBaixar, vbtnChecked_Documentos, vFilial, vNumero_Proposta_trans, vNumero_Documento, vbtnConsultar As String

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
                    vFilial = pc_db.Fieldt("vFilial")
                    vbtnChecked_Documentos = pc_db.Fieldt("vbtnChecked_Documentos")
                    vNumero_Documento = pc_db.Fieldt("vNumero_Documento")
                    vNumero_Proposta_trans = pc_db.Fieldt("vNumero_Proposta_trans")
					vbtnConsultar = pc_db.Fieldt("vbtnConsultar")
					vbtnBaixar = pc_db.Fieldt("vbtnBaixar")
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Controle_de_Documentos.test_KIPREV_Controle_de_Documentos.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
