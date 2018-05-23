'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 27/02/2018 16:47:43 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Web Browser
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Dados de Localização
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Dados_de_Localizacao
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Dados_de_Localizacao
    Public Class test_KIPREV_Dados_de_Localizacao
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            'If p_IsLoop Then ' IS LOOP ALL TESTS
								'If p_OrdemTest > 1 Then
								'	Test.SendKey("{F5}")
								'End If
                            'End If
                            Test.WaitExist("183,101,80,18;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227173458.bmp", typeIdentification.leantest)
                            Test.TestLog("Acessar a funcionalidade Dados de Localização", "Sistema deve apresentar a funcionalidade Dados de Localização", "Sistema apresentou a funcionalidade Dados de Localização com sucesso", typelog.Passed)
							''*******************************************************************************************************************************************************
							'Dados de localização
							Test.Set("65,124,170,22;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227160156.bmp", vTipo_de_Localizacao,"", typeIdentification.leantest) 'type Tipo de Localização
							Test.SendKey("{TAB}")
							Test.Set("267,125,128,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227160452.bmp", vDescricao,"", typeIdentification.leantest) 'type Descrição
								Test.SendKey("{TAB}")
							Test.Set("69,148,64,17;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227160619.bmp", vCEP,"", typeIdentification.leantest) 'type CEP
							Test.Wait(200)
							Test.Set("256,148,215,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227160736.bmp", vTipo_Lograduro,"", typeIdentification.leantest) 'type Tipo Lograduro
							Test.Wait(200)
							Test.Set("502,144,114,22;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227160924.bmp", vLogradouro,"", typeIdentification.leantest) 'type Logradouro
							Test.Wait(200)
							Test.Set("48,165,106,19;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227161000.bmp", vNumero,"", typeIdentification.leantest) 'type Número
							Test.Click("183,167,36,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227161048.bmp", vSem_Numero, typeIdentification.leantest) 'checked Sem Número
							Test.Set("279,166,97,18;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227161130.bmp", vComplemento,"", typeIdentification.leantest) 'type Complemento
							Test.Wait(200)
							Test.Set("512,168,91,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227161158.bmp", vBairro,"", typeIdentification.leantest) 'type Bairro
							Test.Wait(200)
							Test.Set("67,187,60,18;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227161447.bmp", vPais,"", typeIdentification.leantest) 'type Pais
							Test.SendKey("{TAB}")
							Test.Set("293,184,70,21;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227165501.bmp", vUF,"", typeIdentification.leantest) 'type UF
							Test.SendKey("{TAB}")
							Test.Set("507,185,82,20;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227161707.bmp", vCidade,"", typeIdentification.leantest) 'type Cidade
							Test.SendKey("{TAB}")
							''*******************************************************************************************************************************************************
							'Correspondencia
							Test.Click("702,212,18,45;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227162110.bmp", vCorrespondencia, typeIdentification.leantest) 'checked Correspondencia
							''*******************************************************************************************************************************************************
							'Contatos
							Test.Set("59,303,125,25;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180228100254.bmp", vTipo_Telefone,"", typeIdentification.leantest) 'type Tipo Telefone
							Test.Wait(200)
							Test.Set("190,305,102,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227163734.bmp", vLocal_Telefone,"", typeIdentification.leantest) 'type Local Telefone
							Test.Wait(200)
							Test.Set("298,287,61,55;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227163827.bmp", vCod_Area,"", typeIdentification.leantest) 'type Cod Area
							Test.Wait(200)
							Test.Set("340,286,156,55;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227163911.bmp", vTelefone_3,"", typeIdentification.leantest) 'type Telefone
							Test.Wait(200)
							Test.Set("473,286,61,57;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227163943.bmp", vObservacao,"", typeIdentification.leantest) 'type Observação
							Test.Click("685,288,18,54;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227162516.bmp", vPrincipal, typeIdentification.leantest) 'checked Principal
							''*******************************************************************************************************************************************************
							'Email
							Test.Set("287,370,43,55;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227162724.bmp", vEmail,"", typeIdentification.leantest) 'type Email
							Test.Set("599,368,71,57;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227162943.bmp", vTipo_de_Email,"", typeIdentification.leantest) 'type Tipo de Email
							Test.Click("50,427,40,14;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227163129.bmp", vCorrespondencia_Beneficiario, typeIdentification.leantest) 'checked Correspondencia Beneficiario
							Test.Click("50,443,41,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227163236.bmp", vCorrespondencia_Email, typeIdentification.leantest) 'checked Correspondencia Email
							''*******************************************************************************************************************************************************
							'Botoes de Correspondencia
							If CBool(vbtnAlterar) Then
								Test.TestLog("Clicar em Alterar", "Clicar em Alterar e verificar o resultado esperado", "Clique em Alterar com sucesso", typelog.Passed)
								Test.Click("634,434,89,18;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227164132.bmp", vbtnAlterar, typeIdentification.leantest) 'click Alterar
								'Test.TestLog("Resultado após clique em Alterar", "Resultado após clique em Alterar", "Resultado verificado com sucesso", typelog.Passed)
							End If
							If CBool(vbtnCorrespondencia_5) Then
								Test.TestLog("Clicar em Correspondencia", "Clicar em Correspondencia e verificar o resultado esperado", "Clique em Correspondencia com sucesso", typelog.Passed)
								Test.Click("438,434,101,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227164332.bmp", vbtnCorrespondencia_5, typeIdentification.leantest) 'click Correspondencia
								'Test.TestLog("Resultado após clique em Correspondencia", "Resultado após clique em Correspondencia", "Resultado verificado com sucesso", typelog.Passed)
							End If
							''*******************************************************************************************************************************************************
							'Botao de finalizar
							If CBool(vbtnSalvar) Then
								Test.TestLog("Clicar em Salvar", "Clicar em Salvar e verificar o resultado esperado", "Clique em Salvar com sucesso", typelog.Passed)
								Test.Click("752,493,22,26;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227164400.bmp", vbtnSalvar, typeIdentification.leantest) 'click Salvar
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
				HandlerError("test_KIPREV_Dados_de_Localizacao.test_KIPREV_Dados_de_Localizacao.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vTipo_de_Localizacao, vDescricao, vCEP, vTipo_Lograduro, vLogradouro, vNumero, vSem_Numero,vComplemento, vBairro, vPais, vUF, vCidade, vCorrespondencia,vTipo_Telefone, vLocal_Telefone, vCod_Area, vTelefone_3, vObservacao, vPrincipal,vEmail, vTipo_de_Email, vCorrespondencia_Beneficiario,vCorrespondencia_Email,vbtnAlterar,vbtnCorrespondencia_5,vbtnSalvar As String

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
                    vTipo_de_Localizacao = pc_db.Fieldt("vTipo_de_Localizacao")
					vDescricao = pc_db.Fieldt("vDescricao")
					vCEP = pc_db.Fieldt("vCEP")
					vTipo_Lograduro = pc_db.Fieldt("vTipo_Lograduro")
					vLogradouro = pc_db.Fieldt("vLogradouro")
					vNumero = pc_db.Fieldt("vNumero")
					vSem_Numero = pc_db.Fieldt("vSem_Numero")
					vComplemento = pc_db.Fieldt("vComplemento")
					vBairro = pc_db.Fieldt("vBairro")
					vPais = pc_db.Fieldt("vPais")
					vUF = pc_db.Fieldt("vUF")
					vCidade = pc_db.Fieldt("vCidade")
					vCorrespondencia = pc_db.Fieldt("vCorrespondencia")
					vTipo_Telefone = pc_db.Fieldt("vTipo_Telefone")
					vLocal_Telefone = pc_db.Fieldt("vLocal_Telefone")
					vCod_Area = pc_db.Fieldt("vCod_Area")
					vTelefone_3 = pc_db.Fieldt("vTelefone_3")
					vObservacao = pc_db.Fieldt("vObservacao")
					vPrincipal = pc_db.Fieldt("vPrincipal")
					vEmail = pc_db.Fieldt("vEmail")
					vTipo_de_Email = pc_db.Fieldt("vTipo_de_Email")
					vCorrespondencia_Beneficiario = pc_db.Fieldt("vCorrespondencia_Beneficiario")
					vCorrespondencia_Email = pc_db.Fieldt("vCorrespondencia_Email")
					vbtnAlterar = pc_db.Fieldt("vbtnAlterar")
					vbtnCorrespondencia_5 = pc_db.Fieldt("vbtnCorrespondencia_5")
					vbtnSalvar = pc_db.Fieldt("vbtnSalvar")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Dados_de_Localizacao.test_KIPREV_Dados_de_Localizacao.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
