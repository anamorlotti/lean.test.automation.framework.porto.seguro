'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 09/03/2018 09:10:38 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Client Server (Desktop)
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Painel de Controle
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Painel_de_Controle
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Painel_de_Controle
    Public Class test_KIPREV_Painel_de_Controle
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            ''*******************************************************************************************************************************************************
							'Lista de Pessoas - Empresas
							If CBool(vbtnConsulta_por_Nome) Then
								Test.TestLog("Clicar em Consulta por Nome", "Clicar em Consulta por Nome e verificar o resultado esperado", "Clique em Consulta por Nome com sucesso", typelog.Passed)
								Test.Click("117,131,107,14;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309085604.bmp", vbtnConsulta_por_Nome, typeIdentification.leantest) 'click Consulta por Nome
								'Test.TestLog("Resultado após clique em Consulta por Nome", "Resultado após clique em Consulta por Nome", "Resultado verificado com sucesso", typelog.Passed)
							End If
							If CBool(vbtnConsulta_por_Negocio) Then
								Test.TestLog("Clicar em Consulta por Negócio", "Clicar em Consulta por Negócio e verificar o resultado esperado", "Clique em Consulta por Negócio com sucesso", typelog.Passed)
								Test.Click("241,134,122,12;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309085633.bmp", vbtnConsulta_por_Negocio, typeIdentification.leantest) 'click Consulta por Negócio
								'Test.TestLog("Resultado após clique em Consulta por Negócio", "Resultado após clique em Consulta por Negócio", "Resultado verificado com sucesso", typelog.Passed)
							End If
							If CBool(vbtnConsulta_por_Identificacao) Then
								Test.TestLog("Clicar em Consulta por Identificação", "Clicar em Consulta por Identificação e verificar o resultado esperado", "Clique em Consulta por Identificação com sucesso", typelog.Passed)
								Test.Click("380,134,126,18;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309085811.bmp", vbtnConsulta_por_Identificacao, typeIdentification.leantest) 'click Consulta por Identificação
								'Test.TestLog("Resultado após clique em Consulta por Identificação", "Resultado após clique em Consulta por Identificação", "Resultado verificado com sucesso", typelog.Passed)
							End If
							If CBool(vbtnConsulta_por_Agencia) Then
								Test.TestLog("Clicar em Consulta por Agencia", "Clicar em Consulta por Agencia e verificar o resultado esperado", "Clique em Consulta por Agencia com sucesso", typelog.Passed)
								Test.Click("529,132,119,15;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309085739.bmp", vbtnConsulta_por_Agencia, typeIdentification.leantest) 'click Consulta por Agencia
								'Test.TestLog("Resultado após clique em Consulta por Agencia", "Resultado após clique em Consulta por Agencia", "Resultado verificado com sucesso", typelog.Passed)
							End If
							''*******************************************************************************************************************************************************
							'Consulta por Nome
							'Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
							Test.Set("135,167,224,32;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309090033.bmp", vFiltrar_por,"", typeIdentification.leantest) 'type Filtrar por
							Test.SendKey("{Tab}") 'SendKey Tab
							If CBool(vbtnAceitar) Then
								Test.TestLog("Clicar em Aceitar", "Clicar em Aceitar e verificar o resultado esperado", "Clique em Aceitar com sucesso", typelog.Passed)
								Test.Click("311,441,84,19;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309090255.bmp", vbtnAceitar, typeIdentification.leantest) 'click Aceitar
								'Test.TestLog("Resultado após clique em Aceitar", "Resultado após clique em Aceitar", "Resultado verificado com sucesso", typelog.Passed)
							End If
							''*******************************************************************************************************************************************************
							'Painel de Controle
							Test.WaitExist("56,94,42,72;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309090416.bmp", typeIdentification.leantest) 'WaitExist 
							Dim saldo As String 
							Do
								Test.DoubleClick("121,447,115,60;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180316162134.bmp", True, typeIdentification.leantest) 'Dois cliques no certificado
								Test.Wait(4000)
								Test.SendKey("{SELECT}")
								Test.ClipBoardClear
	                            Test.SendKey("^c")
	                            saldo = Test.ClipBoardGetText
	                            Test.SetValueOutput("vSaldo_out", saldo)
	                            
	                        Loop While String.IsNullOrEmpty(saldo) 
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
				HandlerError("test_KIPREV_Painel_de_Controle.test_KIPREV_Painel_de_Controle.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vSaldo_out, vbtnConsulta_por_Nome,vbtnConsulta_por_Negocio,vbtnConsulta_por_Identificacao,vbtnConsulta_por_Agencia,vFiltrar_por, vbtnAceitar As String

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
                    vbtnConsulta_por_Nome = pc_db.Fieldt("vbtnConsulta_por_Nome")
					vbtnConsulta_por_Negocio = pc_db.Fieldt("vbtnConsulta_por_Negocio")
					vbtnConsulta_por_Identificacao = pc_db.Fieldt("vbtnConsulta_por_Identificacao")
					vbtnConsulta_por_Agencia = pc_db.Fieldt("vbtnConsulta_por_Agencia")
					vFiltrar_por = pc_db.Fieldt("vFiltrar_por")
					vbtnAceitar = pc_db.Fieldt("vbtnAceitar")
					vSaldo_out = pc_db.Fieldt("vSaldo_out")
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Painel_de_Controle.test_KIPREV_Painel_de_Controle.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
