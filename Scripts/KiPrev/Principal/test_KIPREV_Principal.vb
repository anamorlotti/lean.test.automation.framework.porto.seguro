'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 26/02/2018 17:27:30 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Web Browser
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Principal
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Principal
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Principal
    Public Class test_KIPREV_Principal
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
							Test.Click("1,22,84,24;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180226170240.bmp", vbtnRADM_KIPREV, typeIdentification.leantest) 'click RADM KIPREV
							
							'***********************************************************************************************************************************************************************************
							'Controles de Documentos
							If CBool(vbtnAdministrativo) Then
							Test.Click("18,192,117,35;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180312153010.bmp", vbtnAdministrativo, typeidentification.leanTest)'click Admintratvo
							End If
							If CBool(vbtnFerramentas) Then
							Test.Click("211,332,104,37;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180312153340.bmp", vbtnFerramentas, typeidentification.leanTest)'click Ferramentas
							End If
							If CBool(vbtnControle_de_Documentos1) Then
							Test.Click("423,450,143,21;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180312153614.bmp", vbtnControle_de_Documentos1, typeidentification.leanTest)'click Controle_de_Documentos1
							End If
							If CBool(vbtnControle_de_Documentos2) Then
							Test.Click("713,455,143,56;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180312153918.bmp", vbtnControle_de_Documentos2, typeidentification.leanTest)'click Controle_de_Documentos2
							End If
											
							''*******************************************************************************************************************************************************
							'Adesao Vida Prev
							If CBool(vbtnAdesao) Then
								Test.Click("19,238,37,34;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180226170852.bmp", vbtnAdesao, typeIdentification.leantest) 'click Adesão
								''*******************************************************************************************************************************************************
							'Adesao PF
								If CBool(vbtnAdesao_PF) Then
								Test.Click("208,274,150,11;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180226171049.bmp", vbtnAdesao_PF, typeIdentification.leantest) 'click Adesao PF
								''*******************************************************************************************************************************************************
							'Adesao Vida Prev
									If CBool(vbtnAdesao_Vida_Prev) Then
									Test.Wait(500)
									Test.TestLog("Acessar a funcionalidade adesão vida prev.", "Sistema deve apresentar a funcionalidade adesão vida prev.", "Sistema apresentou a funcionalidade Principal com sucesso", typelog.Passed)								
                         			Test.Click("384,273,97,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180226171152.bmp", vbtnAdesao_Vida_Prev, typeIdentification.leantest) 'click Adesao Vida Prev
									End If
								End If
							End If
							''*******************************************************************************************************************************************************
							'Propostas
							If CBool(vbtnPropostas) Then
								Test.Click("209,321,59,17;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313105258.bmp", vbtnPropostas, typeIdentification.leantest) 'click Propostas
							End If
							''***************************************************************************************************************************************************
							'Adesao Consulta e Liberação de Proposta Pendente
							If CBool(vbtnConsulta_Liberacao_Proposta_Pendente) Then
								Test.TestLog("Acessar a funcionalidade Consulta e Liberação de Proposta Pendente", "Sistema deve apresentar a funcionalidade Consulta e Liberação de Proposta Pendente", "Sistema apresentou a funcionalidade Consulta e Liberação de Proposta Pendente com sucesso", typelog.Passed)								
                         		Test.Click("380,325,253,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180313105549.bmp", vbtnConsulta_Liberacao_Proposta_Pendente, typeIdentification.leantest) 'click Consulta e Liberação de Proposta Pendente
							End If
							''*******************************************************************************************************************************************************
							'Atendimento
							If CBool(vbtnAtendimento) Then
								Test.TestLog("Acessar a funcionalidade atendimento", "Sistema deve apresentar a funcionalidade atendimento", "Sistema apresentou a funcionalidade atendimento com sucesso", typelog.Passed)		'click atendimento						
                         		Test.Click("54,593,74,21;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180308153848.bmp", vbtnAtendimento, typeIdentification.leantest)
							End If
							''*******************************************************************************************************************************************************
							'Solicitação de aporte
							If CBool(vbtnArrecadacao) Then
								'Test.TestLog("Acessar a funcionalidade Arrecadação", "Sistema deve apresentar a funcionalidade Arrecadação", "Sistema apresentou a funcionalidade Arrecadação com sucesso", typelog.Passed)		'click Arrecadação						
                         		Test.Click("19,310,107,34;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309095459.bmp", vbtnArrecadacao, typeIdentification.leantest)
							End If
							If CBool(vbtnAporte) Then
								'Test.MouseOver("206,312,136,20;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309103405.bmp", vbtnAporte)
								'Test.TestLog("Acessar a funcionalidade atendimento", "Sistema deve apresentar a funcionalidade atendimento", "Sistema apresentou a funcionalidade atendimento com sucesso", typelog.Passed)		'click aporte						
                         		Test.Click("206,464,45,17;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309101828.bmp", vbtnAporte, typeIdentification.leantest)
							End If
							If CBool(vbtnSolicitacao_Aporte_PF) Then
								Test.TestLog("Acessar a funcionalidade Solicitação de Aporte - PF", "Sistema deve apresentar a funcionalidade Solicitação de Aporte - PF", "Sistema deve apresentar a funcionalidade Solicitação de Aporte - PF com sucesso", typelog.Passed)		'click Solicitação de Aporte - PF				
								Test.Click("394,469,200,12;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309100010.bmp", vbtnSolicitacao_Aporte_PF, typeIdentification.leantest)
                         	End If
                         	If CBool(vbtnSolicitacao_Aporte_PJ) Then
                         		Test.TestLog("Acessar a funcionalidade Solicitação de Aporte - PJ", "Sistema deve apresentar a funcionalidade Solicitação de Aporte - PJ", "Sistema deve apresentar a funcionalidade Solicitação de Aporte - PJ com sucesso", typelog.Passed)		'click Solicitação de Aporte - PJ
                         		Test.Click("400,487,191,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309100622.bmp", vbtnSolicitacao_Aporte_PJ, typeIdentification.leantest)
							End If	
								
							''*******************************************************************************************************************************************************
							'Solicitação de Resgate
                         	If CBool(vbtnReserva) Then
                  				Test.Click("17,350,96,22;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309164954.bmp", vbtnReserva, TypeIdentification.leanTest) '	Click Reserva
                        	End If
   							If CBool(vbtnResgate) Then
   								Test.Click("209,476,50,23;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309164721.bmp", vbtnResgate, TypeIdentification.leanTest) 'Click Resgate
   							End If
   							If CBool(vbtnSolicitaca_Resgate) Then
                         		Test.Click("512,483,137,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180309165310.bmp", vbtnSolicitaca_Resgate, TypeIdentification.leanTest) 'Click Solicitação De Resgate
                         		Test.TestLog("Acessar a funcionalidade Solicitação de Resgate", "Sistema deve apresentar a funcionalidade Solicitação de Resgate", "Sistema deve apresentar a funcionalidade Solicitação de Resgate com sucesso", typelog.Passed)		'click Solicitação de Resgate
                    		End If
                    		If CBool (vbtnPortabilidade) Then
                    			Test.Click("205,354,75,19;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180314171106.bmp", vbtnPortabilidade, Typeidentification.leanTest)'click em Portabilidade
                    		End If
                    		If CBool(vbtnSolicitacao_Portabilidade_Entrada)Then
                    			Test.Click("506,379,227,19;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180314171644.bmp", vbtnSolicitacao_Portabilidade_Entrada, TypeIdentification.leanTest) 'click em Solicitação de Portabilidade de Entrada
                    			Test.TestLog("Acessar a funcionalidade Solicitação de Portabilidade de Entrada", "Sistema deve apresentar a funcionalidade Solicitação de Portabilidade de Entrada", "Sistema deve apresentar a funcionalidade Solicitação de Resgate com sucesso", typelog.Passed)		
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
				HandlerError("test_KIPREV_Principal.test_KIPREV_Principal.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vbtnSolicitacao_Portabilidade_Entrada, vbtnPortabilidade, vbtnConsulta_Liberacao_Proposta_Pendente, vbtnPropostas, vbtnAdministrativo, vbtnFerramentas,vbtnControle_de_Documentos1, vbtnControle_de_Documentos2, vbtnResgate, vbtnSolicitaca_Resgate, vbtnReserva, vbtnArrecadacao, vbtnAporte, vbtnSolicitacao_Aporte_PF,vbtnSolicitacao_Aporte_PJ,
        				vbtnRADM_KIPREV,vbtnAdesao,vbtnAdesao_PF,vbtnAdesao_Vida_Prev, vbtnAtendimento As String

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
					
					vbtnSolicitacao_Portabilidade_Entrada = pc_db.Fieldt("vbtnSolicitacao_Portabilidade_Entrada")
					vbtnPortabilidade = pc_db.Fieldt("vbtnPortabilidade")
					vbtnReserva = pc_db.Fieldt("vbtnReserva")
				   	vbtnSolicitaca_Resgate = pc_db.Fieldt("vbtnSolicitaca_Resgate")
					vbtnResgate = pc_db.Fieldt("vbtnResgate")
				   	vbtnArrecadacao = pc_db.Fieldt("vbtnArrecadacao")
                    vbtnAporte = pc_db.Fieldt ("vbtnAporte")
					vbtnSolicitacao_Aporte_PF = pc_db.Fieldt("vbtnSolicitacao_Aporte_PF")
					vbtnSolicitacao_Aporte_PJ = pc_db.Fieldt("vbtnSolicitacao_Aporte_PJ")
                    vbtnRADM_KIPREV = pc_db.Fieldt("vbtnRADM_KIPREV")
                    vbtnAtendimento = pc_db.Fieldt ("vbtnAtendimento")
					vbtnAdesao = pc_db.Fieldt("vbtnAdesao")
					vbtnAdesao_PF = pc_db.Fieldt("vbtnAdesao_PF")
					vbtnAdesao_Vida_Prev = pc_db.Fieldt("vbtnAdesao_Vida_Prev")
					vbtnArrecadacao = pc_db.Fieldt("vbtnArrecadacao")
					vbtnAporte = pc_db.Fieldt("vbtnAporte")
					vbtnAdministrativo = pc_db.Fieldt("vbtnAdministrativo")
					vbtnFerramentas = pc_db.Fieldt("vbtnFerramentas")
					vbtnControle_de_Documentos1 = pc_db.Fieldt("vbtnControle_de_Documentos1")
					vbtnControle_de_Documentos2 = pc_db.Fieldt("vbtnControle_de_Documentos2")
					vbtnPropostas = pc_db.Fieldt("vbtnPropostas")	
					vbtnConsulta_Liberacao_Proposta_Pendente = pc_db.Fieldt("vbtnConsulta_Liberacao_Proposta_Pendente")
					
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Principal.test_KIPREV_Principal.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
