'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 19/04/2018 15:54:52 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... Orquestrador Poposta
'Functionality:... Localizar Proposta no ESEG
'Master Test:..... No Defined
'TableTest:....... test_Orquestrador_Poposta_Localizar_Proposta_no_ESEG
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Orquestrador_Poposta_Localizar_Proposta_no_ESEG
    Public Class test_Orquestrador_Poposta_Localizar_Proposta_no_ESEG
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "http://osbhmlcrp.portoseguro.brasil:80/OrquestradorIntegrationService/v1/osb/proxy/OrquestradorPropostaPDCServicePSV1_0"'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters Orquestrador Poposta
							vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope Orquestrador Poposta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_Cd_filial", vv1_Cd_filial,"", vXMLRequestEnvelope) 'type value in element xmlv1:Cd_filial
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_Cd_produto", vv1_Cd_produto,"", vXMLRequestEnvelope) 'type value in element xmlv1:Cd_produto
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_Cd_proposta", vv1_Cd_proposta,"", vXMLRequestEnvelope) 'type value in element xmlv1:Cd_proposta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_succod", vv1_succod,"", vXMLRequestEnvelope) 'type value in element xmlv1:succod
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_ramcod", vv1_ramcod,"", vXMLRequestEnvelope) 'type value in element xmlv1:ramcod
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_aplnumdig", vv1_aplnumdig,"", vXMLRequestEnvelope) 'type value in element xmlv1:aplnumdig
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_prporg", vv1_prporg,"", vXMLRequestEnvelope) 'type value in element xmlv1:prporg
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_prpnumdig", vv1_prpnumdig,"", vXMLRequestEnvelope) 'type value in element xmlv1:prpnumdig
							Test.TestLog("Consumir o serviço Localizar Proposta no ESEG", "O serviço Localizar Proposta no ESEG deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço Localizar Proposta no ESEG foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService Orquestrador Poposta
							If p_xmlSoapLeanTest.StatusCode = 200 Then
								Test.TestLog("Verificar o retorno do serviço Localizar Proposta no ESEG", "O serviço Localizar Proposta no ESEG retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço Localizar Proposta no ESEG foi consumido com sucesso", typelog.Passed, , False)
							Else
							Test.TestLog("Verificar o retorno do serviço Localizar Proposta no ESEG", "O serviço Localizar Proposta no ESEG retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço Localizar Proposta no ESEG não foi consumido com sucesso", typelog.Failed, , False)
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
				HandlerError("test_Orquestrador_Poposta_Localizar_Proposta_no_ESEG.test_Orquestrador_Poposta_Localizar_Proposta_no_ESEG.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vv1_Cd_filial, vv1_Cd_produto, vv1_Cd_proposta, vv1_succod, vv1_ramcod, vv1_aplnumdig, vv1_prporg, vv1_prpnumdig, vIsOpenSystem As String

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
					vv1_Cd_filial = pc_db.Fieldt("vv1_Cd_filial")
					vv1_Cd_produto = pc_db.Fieldt("vv1_Cd_produto")
					vv1_Cd_proposta = pc_db.Fieldt("vv1_Cd_proposta")
					vv1_succod = pc_db.Fieldt("vv1_succod")
					vv1_ramcod = pc_db.Fieldt("vv1_ramcod")
					vv1_aplnumdig = pc_db.Fieldt("vv1_aplnumdig")
					vv1_prporg = pc_db.Fieldt("vv1_prporg")
					vv1_prpnumdig = pc_db.Fieldt("vv1_prpnumdig")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Orquestrador_Poposta_Localizar_Proposta_no_ESEG.test_Orquestrador_Poposta_Localizar_Proposta_no_ESEG.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
