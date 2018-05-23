'***************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 27/02/2018 10:32:01 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Web Browser
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Adesao_Dados Basicos
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Dados_Basicos
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Dados_Basicos
    Public Class test_KIPREV_Dados_Basicos
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                    	Try
                            Test.WaitExist("14,72,132,47;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227140042.bmp", typeidentification.leanTest)         		  
							''*******************************************************************************************************************************************************
							'Proposta só será individual
							Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
							Test.Click("136,210,59,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227092323.bmp", vPortabilidade_de_Entrada, typeIdentification.leantest) 'checked Portabilidade de Entrada
							Test.Set("206,211,135,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227092453.bmp", vCertificado,"", typeIdentification.leantest) 'type Certificado
							Test.Click("495,206,50,23;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227092556.bmp", vContem_info_restrita, typeIdentification.leantest) 'checked Contem info restrita
							''*******************************************************************************************************************************************************
							'Informações do Cliente
							Test.Set("73,287,124,23;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180306083919.bmp", vCPF,"", typeIdentification.leantest) 'type CPF
       						 Test.SendKey("{ENTER}")
                            Test.Wait(500)
                            Test.SendKey("{ENTER}")
							Test.Click("283,290,38,15;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227093046.bmp", vCPF_proprio, typeIdentification.leantest) 'checked CPF proprio
							'sexo****************************************************************************************************************************************************
							Test.set("329,282,71,32;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227093859.bmp", vSexo ,, typeIdentification.leantest) 'click sexo
							Test.Set("543,284,70,26;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227094642.bmp", vData_de_Nascimento,"", typeIdentification.leantest) 'type Data de Nascimento
							''*******************************************************************************************************************************************************
							'Documento adicional se for um novo cadastro
							If CBool(vbtnDocumentos_Adicionais) Then
								Test.Click("736,285,28,19;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227095406.bmp", vbtnDocumentos_Adicionais, typeIdentification.leantest) 'click Documentos Adicionais
								Test.Click("202,127,79,102;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227141641.bmp", True, typeidentification.leanTest)
								Test.Click("13,166,83,22;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227141324.bmp", true, typeIdentification.leantest) 'click Tipo de Documento
								'selecao do tipo de documento
								Test.SendKey(VTipoDocumento)
								Test.SendKey("{Tab}")
								Test.SendKey(vDocumento)
								Test.TestLog("Clicar em OK", "Clicar em OK e verificar o resultado esperado", "Clique em OK com sucesso", typelog.Passed)
								Test.Click("340,384,16,9;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227100202.bmp", true, typeIdentification.leantest) 'click OK
								Test.TestLog("Resultado após clique em OK", "Resultado após clique em OK", "Resultado verificado com sucesso", typelog.Passed)
							End If
							'fim documentos adicionais*******************************************************************************************************************************
							''*******************************************************************************************************************************************************
							'Dados Cadastrais
							Test.Set("16,334,203,23;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227163545.bmp", vNome_Completo,"", typeIdentification.leantest) 'type Nome Completo
							Test.Select("534,355,91,18;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227151458.bmp", vPep, typeidentification.leanTest) 'select pep
							Test.Click("749,354,20,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227101552.bmp", vReside_no_Brasil, typeIdentification.leantest) 'checked Reside no Brasil
							Test.Select("71,376,59,14;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227101949.bmp", vEstado_Civil, typeIdentification.leantest) 'type Estado Civil
							Test.Set("46,391,139,18;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227153458.bmp", vProfissao,"", typeIdentification.leantest) 'type Profissao
							Test.Set("72,410,69,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227102445.bmp", vRenda_Mensal,"", typeIdentification.leantest) 'type Renda Mensal
							Test.Set("210,409,62,19;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227102516.bmp", vPatrimonio,"", typeIdentification.leantest) 'type Patrimonio
							Test.Click("618,391,60,18;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227102634.bmp", vDispensa_NF, typeIdentification.leantest) 'checked Dispensa NF
							If CBool(vbtnSalvar) Then
								Test.WaitExist("618,391,60,18;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227102634.bmp", typeIdentification.leantest) 'checked Dispensa NF
								Test.TestLog("Clicar em Salvar", "Clicar em Salvar e verificar o resultado esperado", "Clique em Salvar com sucesso", typelog.Passed)
								Test.Click("747,493,32,24;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180227102922.bmp", vbtnSalvar, typeIdentification.leantest) 'click Salvar
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
				HandlerError("test_KIPREV_Adesao_Dados_Basicos.test_KIPREV_Adesao_Dados_Basicos.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vPortabilidade_de_Entrada,vCertificado, vContem_info_restrita,vCPF, vCPF_proprio, vSexo,vData_de_Nascimento, 
        	vbtnDocumentos_Adicionais, VTipoDocumento,vDocumento, vNome_Completo, vPep, vReside_no_Brasil, vEstado_Civil, 
        	vProfissao, vRenda_Mensal, vPatrimonio, vDispensa_NF,vbtnSalvar As String

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
                    vPortabilidade_de_Entrada = pc_db.Fieldt("vPortabilidade_de_Entrada")
					vCertificado = pc_db.Fieldt("vCertificado")
					vContem_info_restrita = pc_db.Fieldt("vContem_info_restrita")
					vCPF = pc_db.Fieldt("vCPF")
					vCPF_proprio = pc_db.Fieldt("vCPF_proprio")
					vSexo = pc_db.Fieldt("vSexo")
					vData_de_Nascimento = pc_db.Fieldt("vData_de_Nascimento")
					vbtnDocumentos_Adicionais = pc_db.Fieldt("vbtnDocumentos_Adicionais")
					VTipoDocumento = pc_db.Fieldt("VTipoDocumento")
					vDocumento = pc_db.Fieldt("vDocumento")
					vNome_Completo = pc_db.Fieldt("vNome_Completo")
					vPep = pc_db.Fieldt("vPep")
					vReside_no_Brasil = pc_db.Fieldt("vReside_no_Brasil")
					vEstado_Civil = pc_db.Fieldt("vEstado_Civil")
					vProfissao = pc_db.Fieldt("vProfissao")
					vRenda_Mensal = pc_db.Fieldt("vRenda_Mensal")
					vPatrimonio = pc_db.Fieldt("vPatrimonio")
					vDispensa_NF = pc_db.Fieldt("vDispensa_NF")
					vbtnSalvar = pc_db.Fieldt("vbtnSalvar")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Adesao_Dados_Basicos.test_KIPREV_Adesao_Dados_Basicos.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
