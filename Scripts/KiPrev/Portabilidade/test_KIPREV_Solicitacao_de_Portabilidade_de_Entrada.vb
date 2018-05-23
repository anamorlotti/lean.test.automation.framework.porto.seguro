'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 15/03/2018 11:47:02 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to Client Server (Desktop)
'Environment:..... Automation Project
'Application...... KIPREV
'Functionality:... Solicitaçao de Portabilidade de Entrada
'Master Test:..... No Defined
'TableTest:....... test_KIPREV_Solicitacao_de_Portabilidade_de_Entrada
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_KIPREV_Solicitacao_de_Portabilidade_de_Entrada
    Public Class test_KIPREV_Solicitacao_de_Portabilidade_de_Entrada
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            ''*******************************************************************************************************************************************************
                            'Identificação
                           	Test.WaitExist("616,118,162,18;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315110305.bmp", typeIdentification.leantest) 'select Tipo Portabilidade
							Test.TestLog("Informar valores", "Sistema deve solicitar os valores conforme critérios de entrada", "Sistema permitiu a inclusão de valores com sucesso", typelog.Passed)
							
							Test.Select("616,118,162,18;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315110305.bmp", vTipo_Portabilidade, typeIdentification.leantest) 'select Tipo Portabilidade
							Test.SendKey("{Enter}")	
							Test.SendKey("{TAB}")	
							Test.Set("30,194,118,24;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315110500.bmp", vNumero_Proposta_trans,"", typeIdentification.leantest) 'type Conta
							''*******************************************************************************************************************************************************
							'Dados do Plano
							Test.WaitExist("240,199,166,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315113827.bmp",typeIdentification.leantest) 'WaitExist Dt. Prof.
							Test.Set("240,199,166,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315113827.bmp", vDt_Prof,"", typeIdentification.leantest) 'type Dt. Prof.
							Test.Wait(4000)
							Test.DoubleClick("416,193,113,25;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315160048.bmp", True, TypeIdentification.leanTest) ' clck no benficio
							Test.SendKey("{ENTER}")		
							Test.Select("410,221,176,14;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315110915.bmp", vPreenc, typeIdentification.leantest) 'select Preenc.
							Test.Wait(3000)	
							Test.Set("589,216,193,22;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315111014.bmp", vValor_Solicitado,"", typeIdentification.leantest) 'type Valor Solicitado
							Test.Wait(300)	
							Test.Set("240,199,166,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315113827.bmp", vDt_Prof,"", typeIdentification.leantest) 'type Dt. Prof.
							
							'Porcentagem**************************************************************************************************************
							Dim Porcentagem As String() = vPorcentagem.Split(";")
							
							For i= 0 To CInt(vPorcentagem)
								'seleciona a primeira linha
								If i < 1 Then'seleciona a primeira linha
									Test.Click("518,236,101,57;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315111542.bmp", True, typeidentification.leanTest)	
									Test.Set("518,236,101,57;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315111542.bmp", Porcentagem(i),"", typeIdentification.leantest) 'type Porcentagem
							
								Elseif i = 1  Then'seleciona as demais linhas
									Test.Click("518,236,101,57;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315111542.bmp", True, typeidentification.leanTest)	
									Test.SendKey("{DOWN}", i)
									Test.Set("518,236,101,57;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315111542.bmp", Porcentagem(i),"", typeIdentification.leantest) 'type Porcentagem
								End If
							Next
		
							''*******************************************************************************************************************************************************
							'Dados Cedente
							Test.Set("32,329,105,16;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315112015.bmp", vCedente,"", typeIdentification.leantest) 'type Cedente
							Test.Wait(3000)	
							Test.Set("14,344,283,23;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315112202.bmp", vPrc_SUSEP,"", typeIdentification.leantest) 'type Prc. SUSEP
							Test.Wait(3000)	
							Test.Set("304,351,219,17;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315112502.bmp", vCNPJ_Fundo,"", typeIdentification.leantest) 'type CNPJ Fundo
							Test.Wait(3000)	
							Test.Select("529,346,244,22;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315112629.bmp", vReg_Tributario, typeIdentification.leantest) 'select Reg. Tributário
							Test.Wait(3000)	
							Test.Set("32,367,149,19;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315112812.bmp", vProduto,"", typeIdentification.leantest) 'type Produto
							Test.Wait(3000)	
							Test.Set("349,370,59,15;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315112848.bmp", vModalidade,"", typeIdentification.leantest) 'type Modalidade
							Test.Wait(3000)	
							Test.Set("542,371,196,13;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315112939.bmp", vCertificado,"", typeIdentification.leantest) 'type Certificado
							Test.Wait(3000)		
							''*******************************************************************************************************************************************************
							'Dados do Pagamento
							Test.Set("12,401,131,21;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315113134.bmp", vForma_Pagamento,"", typeIdentification.leantest) 'type Forma  Pagamento
							''*******************************************************************************************************************************************************
							'Retorno
							If CBool(vbtnValidar) Then
								Test.TestLog("Clicar em Validar", "Clicar em Validar e verificar o resultado esperado", "Clique em Validar com sucesso", typelog.Passed)
								Test.Click("71,528,37,15;C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\imgCaptured\img_20180315113533.bmp", vbtnValidar, typeIdentification.leantest) 'click Validar
								Test.TestLog("Resultado após clique em Validar", "Resultado após clique em Validar", "Resultado verificado com sucesso", typelog.Passed)
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
				HandlerError("test_KIPREV_Solicitacao_de_Portabilidade_de_Entrada.test_KIPREV_Solicitacao_de_Portabilidade_de_Entrada.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vBenef, vNumero_Proposta_trans, vTipo_Portabilidade,vConta, vDt_Prof, vPreenc,vValor_Solicitado, vPorcentagem, vCedente, vPrc_SUSEP, vCNPJ_Fundo, vReg_Tributario,vProduto, vModalidade, vCertificado, vForma_Pagamento, vbtnValidar As String

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
                    vTipo_Portabilidade = pc_db.Fieldt("vTipo_Portabilidade")
					vConta = pc_db.Fieldt("vConta")
					vDt_Prof = pc_db.Fieldt("vDt_Prof")
					vPreenc = pc_db.Fieldt("vPreenc")
					vValor_Solicitado = pc_db.Fieldt("vValor_Solicitado")
					vPorcentagem = pc_db.Fieldt("vPorcentagem")
					vCedente = pc_db.Fieldt("vCedente")
					vPrc_SUSEP = pc_db.Fieldt("vPrc_SUSEP")
					vCNPJ_Fundo = pc_db.Fieldt("vCNPJ_Fundo")
					vReg_Tributario = pc_db.Fieldt("vReg_Tributario")
					vProduto = pc_db.Fieldt("vProduto")
					vModalidade = pc_db.Fieldt("vModalidade")
					vCertificado = pc_db.Fieldt("vCertificado")
					vForma_Pagamento = pc_db.Fieldt("vForma_Pagamento")
					vbtnValidar = pc_db.Fieldt("vbtnValidar")
					vNumero_Proposta_trans = pc_db.Fieldt("vNumero_Proposta_trans")
					vBenef = pc_db.Fieldt("vBenef")
					
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_KIPREV_Solicitacao_de_Portabilidade_de_Entrada.test_KIPREV_Solicitacao_de_Portabilidade_de_Entrada.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
