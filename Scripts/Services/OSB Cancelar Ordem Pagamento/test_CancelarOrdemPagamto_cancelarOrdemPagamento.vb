'**********************************************************************************************************************************************************************
'Create by LeanTest Automation 3.8 in 05/04/2018 11:38:30 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... CancelarOrdemPagamto
'Functionality:... cancelarOrdemPagamento
'Master Test:..... No Defined
'TableTest:....... test_CancelarOrdemPagamto_cancelarOrdemPagamento
'**********************************************************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_CancelarOrdemPagamto_cancelarOrdemPagamento
    Public Class test_CancelarOrdemPagamto_cancelarOrdemPagamento
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "http://osbquantumhml.portoseguro.brasil:80/Financeiro/CancelarOrdemPagamentoIntegrationInV1.1/Proxy_Services/CancelarOrdemPagamentoIntegrationInV1_1"'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters CancelarOrdemPagamto
							vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope CancelarOrdemPagamto
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcan_codigoEmpresa", vcan_codigoEmpresa,"", vXMLRequestEnvelope) 'type value in element xmlcan:codigoEmpresa
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcan_origem", vcan_origem,"", vXMLRequestEnvelope) 'type value in element xmlcan:origem
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcan_referencia", vcan_referencia,"", vXMLRequestEnvelope) 'type value in element xmlcan:referencia
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcan_docSAP", vcan_docSAP,"", vXMLRequestEnvelope) 'type value in element xmlcan:docSAP
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcan_motivoCancelamento", vcan_motivoCancelamento,"", vXMLRequestEnvelope) 'type value in element xmlcan:motivoCancelamento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcan_descricaoCancelamento", vcan_descricaoCancelamento,"", vXMLRequestEnvelope) 'type value in element xmlcan:descricaoCancelamento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcan_dataCancelamento", vcan_dataCancelamento,"", vXMLRequestEnvelope) 'type value in element xmlcan:dataCancelamento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcan_matricula", vcan_matricula,"", vXMLRequestEnvelope) 'type value in element xmlcan:matricula
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcan_indTributavel", vcan_indTributavel,"", vXMLRequestEnvelope) 'type value in element xmlcan:indTributavel
							Test.TestLog("Consumir o serviço cancelarOrdemPagamento", "O serviço cancelarOrdemPagamento deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço cancelarOrdemPagamento foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService CancelarOrdemPagamto
                            If p_xmlSoapLeanTest.StatusCode = 200 Then
                                Test.TestLog("Verificar o retorno do serviço cancelarOrdemPagamento", "O serviço cancelarOrdemPagamento retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço cancelarOrdemPagamento foi consumido com sucesso", typelog.Passed, , False)
                            Else
                                Test.TestLog("Verificar o retorno do serviço cancelarOrdemPagamento", "O serviço cancelarOrdemPagamento retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço cancelarOrdemPagamento não foi consumido com sucesso", typelog.Failed, , False)
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
		HandlerError("test_CancelarOrdemPagamto_cancelarOrdemPagamento.test_CancelarOrdemPagamto_cancelarOrdemPagamento.Run: " & ex.Message)
		Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
		Return False
            End Try
        End Function

        '****************************************************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vcan_codigoEmpresa, vcan_origem, vcan_referencia, vcan_docSAP, vcan_motivoCancelamento, vcan_descricaoCancelamento, vcan_dataCancelamento, vcan_matricula, vcan_indTributavel, vIsOpenSystem As String

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
                    vXMLParameters = pc_db.Fieldt("vXMLParameters")
					vXMLRequestEnvelope = pc_db.Fieldt("vXMLRequestEnvelope")
					vcan_codigoEmpresa = pc_db.Fieldt("vcan_codigoEmpresa")
					vcan_origem = pc_db.Fieldt("vcan_origem")
					vcan_referencia = pc_db.Fieldt("vcan_referencia")
					vcan_docSAP = pc_db.Fieldt("vcan_docSAP")
					vcan_motivoCancelamento = pc_db.Fieldt("vcan_motivoCancelamento")
					vcan_descricaoCancelamento = pc_db.Fieldt("vcan_descricaoCancelamento")
					vcan_dataCancelamento = pc_db.Fieldt("vcan_dataCancelamento")
					vcan_matricula = pc_db.Fieldt("vcan_matricula")
					vcan_indTributavel = pc_db.Fieldt("vcan_indTributavel")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_CancelarOrdemPagamto_cancelarOrdemPagamento.test_CancelarOrdemPagamto_cancelarOrdemPagamento.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
