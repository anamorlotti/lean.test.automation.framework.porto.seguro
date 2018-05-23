'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 06/03/2018 17:48:49 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Client Server (Desktop)
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Respostas do Cliente
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Respostas_do_Cliente
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Respostas_do_Cliente
    Public Class test_KIPREV_Respostas_do_Cliente
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                        	Test.WaitExist("38,248,53,26;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180306163409.bmp", typeIdentification.leanTest) 'Wait exist select Questionario
                            Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)

                            ''*******************************************************************************************************************************************************
                            'Seleção do tipo de Questionário
                            Select Case vQuestionario
                            	Case "DPSS"
                                    'Test.Click("", True, typeIdentification.leanTest)
                                Case "DPSC"
                                    Test.Click("38,248,53,26;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180306163409.bmp", True, typeIdentification.leanTest) 'select Questionario
                                    Test.Click("39,235,42,17;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310181659.bmp", True, typeIdentification.leanTest)
                            End Select
                            ''*******************************************************************************************************************************************************

                            Dim count As Integer
                            Dim lastCodQuestion As String = Nothing
                            Dim idQuestion As String = Nothing

                            If vQuestionario = "DPSS" Then count = 2 Else count = 13
								Dim flagFirstQuestion As Boolean = True

                            For i = 1 To 20

                                If i <= 2 Then
                                	'Test.Click("62,297,44,60;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310155422.bmp", True, typeIdentification.leanTest)
                                	Test.Click("379,189,60,45;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180523144922.bmp", True, typeIdentification.leanTest)
                                    If i > 1 Then
                                        Test.SendKey("{DOWN}", i - 1)
                                    Else
                                        Test.SendKey("{DOWN}")
                                        Test.SendKey("{UP}")
                                    End If
                                Else
                                    Test.Click("66,306,44,31;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310193648.bmp", True, typeIdentification.leanTest)
                                    If i <= 6 Then
                                        Test.SendKey("{DOWN}", i - 1)
                                    Else
                                        Test.SendKey("{DOWN}", 6)
                                    End If
                                End If

                                Test.SendKey("^c")
                                idQuestion = Test.ClipBoardGetText
                                If lastCodQuestion <> idQuestion Then
                                    lastCodQuestion = idQuestion
                                Else
                                    Exit For
                                End If
                                'Test.Click("479,295,55,61;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310163941.bmp", True, typeIdentification.leanTest)
                                Test.Click("24,348,50,56;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180523151951.bmp", True, typeIdentification.leanTest)
                                Select Case idQuestion
                                    Case 1 'imc
                                        Dim valueQuest1 As String() = v1_IMC.Split(";")
                                        Test.SendKey("1")
                                        Test.SendKey("{Tab}")
                                        Test.SendKey("{Enter}")
                                        Test.SendKey(valueQuest1(0))
                                        Test.Click("111,242,70,31;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310183459.bmp", True, typeIdentification.leanTest)
                                        Test.SendKey("{DOWN}")
                                        Test.SendKey("2")
                                        Test.SendKey("{Tab}")
                                        Test.SendKey("{Enter}")
                                        Test.SendKey(valueQuest1(1))
                                        Test.Click("111,242,70,31;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310183459.bmp", True, typeIdentification.leanTest)
                                    Case 2
                                        Test.SendKey(v2_Fumante_quanto_tempo)
                                        Test.SendKey("{Tab}")
                                    Case 20
                                        Test.SendKey(v20_Ex_fumante_quanto_tempo)
                                        Test.SendKey("{Tab}")
                                    Case 3
                                        Test.SendKey(v3_O_proponente_ira_custear)
                                        Test.SendKey("{Tab}")
                                    Case 4
                                        Test.SendKey(v4_Encontra_se_atualmente_em_boas_condições_de_saude)
                                        Test.SendKey("{Tab}")
                                    Case 5
                                        Test.SendKey(v5_Sofreu_ou_sofre_de_alguma_doenca)
                                        Test.SendKey("{Tab}")
                                    Case 6
                                        Test.SendKey(v6_Tem_ou_teve_alguma_das_doencas_cardiovasculares)
                                        Test.SendKey("{Tab}")
                                    Case 7
                                        Test.SendKey(v7_Ja_realizou_exames_que_tenham_apresentado_resultado)
                                        Test.SendKey("{Tab}")
                                    Case 8
                                        Test.SendKey(v8_Tem_ou_teve_doencas_endocrinas)
                                        Test.SendKey("{Tab}")
                                    Case 9
                                        Test.SendKey(v9_Possui_ou_possuiu_algum_timpo_de_cancer)
                                        Test.SendKey("{Tab}")
                                    Case 10
                                        Test.SendKey(v10_Ja_teve_internado_realizou_tratamento)
                                        Test.SendKey("{Tab}")
                                    Case 11
                                        Test.SendKey(v11_Pratica_esporte_de_forma_profissional)
                                        Test.SendKey("{Tab}")
                                    Case 12
                                        Test.SendKey(v12_E_tripulante_ou_piloto_amador_ou_profissional)
                                        Test.SendKey("{Tab}")
                                    Case 13
                                        Test.SendKey(v13_Possui_seguro_de_vida_em_outra_seguradora)
                                        Test.SendKey("{Tab}")
                                    Case 20
                                        Test.SendKey(v20_Ex_fumante_quanto_tempo)
                                        Test.SendKey("{Tab}")
                                    Case 21
                                    	Select Case  v21_Perfeita_Condicao_de_Saude
                                    		Case 1
                                    			Test.SendKey(v21_Perfeita_Condicao_de_Saude)
		                                        Test.SendKey("{Tab}")
		                                        Test.SendKey("{Enter}")
		                                        Test.SendKey("@Word;5;4")
		                                        Test.Click("111,242,70,31;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180310183459.bmp", True, typeIdentification.leanTest)
                                    	    Case 2
                                    	    	Test.SendKey(v21_Perfeita_Condicao_de_Saude)
		                                        Test.SendKey("{Tab}")
										End Select
                                    Case 22
                                        Test.SendKey(v22_Fumante_quanto_tempo)
                                        Test.SendKey("{Tab}")
                                End Select
                            Next
                            ''*******************************************************************************************************************************************************
                            'Fim do questionario - salvar
                            If CBool(vbtnSalvar) Then
								Test.TestLog("Clicar em Salvar", "Clicar em Salvar e verificar o resultado esperado", "Clique em Salvar com sucesso", typelog.Passed)
								Test.Click("751,494,25,21;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180306173003.bmp", vbtnSalvar, typeIdentification.leantest) 'click Salvar
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
				HandlerError("test_KIPREV_Respostas_do_Cliente.test_KIPREV_Respostas_do_Cliente.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vQuestionario, v21_Perfeita_Condicao_de_Saude, v22_Fumante_quanto_tempo,
            v1_IMC, v2_Fumante_quanto_tempo, v20_Ex_fumante_quanto_tempo, v3_O_proponente_ira_custear,
            v4_Encontra_se_atualmente_em_boas_condições_de_saude, v5_Sofreu_ou_sofre_de_alguma_doenca, v6_Tem_ou_teve_alguma_das_doencas_cardiovasculares,
            v7_Ja_realizou_exames_que_tenham_apresentado_resultado, v8_Tem_ou_teve_doencas_endocrinas, v9_Possui_ou_possuiu_algum_timpo_de_cancer,
            v10_Ja_teve_internado_realizou_tratamento, v11_Pratica_esporte_de_forma_profissional, v12_E_tripulante_ou_piloto_amador_ou_profissional,
            v13_Possui_seguro_de_vida_em_outra_seguradora, vbtnSalvar As String

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
                    
                    vQuestionario = pc_db.Fieldt("vQuestionario")
					v21_Perfeita_Condicao_de_Saude = pc_db.Fieldt("v21_Perfeita_Condicao_de_Saude")
					v22_Fumante_quanto_tempo = pc_db.Fieldt("v22_Fumante_quanto_tempo")
                    v1_IMC = pc_db.Fieldt("v1_IMC")
                    v2_Fumante_quanto_tempo = pc_db.Fieldt("v2_Fumante_quanto_tempo")
                    v20_Ex_fumante_quanto_tempo = pc_db.Fieldt("v20_Ex_fumante_quanto_tempo")
                    v3_O_proponente_ira_custear = pc_db.Fieldt("v3_O_proponente_ira_custear")
					v4_Encontra_se_atualmente_em_boas_condições_de_saude = pc_db.Fieldt("v4_Encontra_se_atualmente_em_boas_condições_de_saude")
					v5_Sofreu_ou_sofre_de_alguma_doenca = pc_db.Fieldt("v5_Sofreu_ou_sofre_de_alguma_doenca")
					v6_Tem_ou_teve_alguma_das_doencas_cardiovasculares = pc_db.Fieldt("v6_Tem_ou_teve_alguma_das_doencas_cardiovasculares")
					v7_Ja_realizou_exames_que_tenham_apresentado_resultado = pc_db.Fieldt("v7_Ja_realizou_exames_que_tenham_apresentado_resultado")
					v8_Tem_ou_teve_doencas_endocrinas = pc_db.Fieldt("v8_Tem_ou_teve_doencas_endocrinas")
					v9_Possui_ou_possuiu_algum_timpo_de_cancer = pc_db.Fieldt("v9_Possui_ou_possuiu_algum_timpo_de_cancer")
					v10_Ja_teve_internado_realizou_tratamento = pc_db.Fieldt("v10_Ja_teve_internado_realizou_tratamento")
					v11_Pratica_esporte_de_forma_profissional = pc_db.Fieldt("v11_Pratica_esporte_de_forma_profissional")
					v12_E_tripulante_ou_piloto_amador_ou_profissional = pc_db.Fieldt("v12_E_tripulante_ou_piloto_amador_ou_profissional")
					v13_Possui_seguro_de_vida_em_outra_seguradora = pc_db.Fieldt("v13_Possui_seguro_de_vida_em_outra_seguradora")
					vbtnSalvar = pc_db.Fieldt("vbtnSalvar")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Respostas_do_Cliente.test_KIPREV_Respostas_do_Cliente.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
