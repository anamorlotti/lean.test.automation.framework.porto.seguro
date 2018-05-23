'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 05/04/2018 15:06:23 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... OSB LocalizaPropESEG
'Functionality:... ConsultaProposta
'Master Test:..... No Defined
'TableTest:....... test_OSB_LocalizaPropESEG_ConsultaProposta
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_OSB_LocalizaPropESEG_ConsultaProposta
    Public Class test_OSB_LocalizaPropESEG_ConsultaProposta
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "http://nt360.portoseguro.brasil/eseg/WebServices/ConsultaMK/consultaPropostaMK.asmx"'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters OSB LocalizaPropESEG
							vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope OSB LocalizaPropESEG
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtem_Cd_filial", vtem_Cd_filial,"", vXMLRequestEnvelope) 'type value in element xmltem:Cd_filial
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtem_Cd_produto", vtem_Cd_produto,"", vXMLRequestEnvelope) 'type value in element xmltem:Cd_produto
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtem_Cd_proposta", vtem_Cd_proposta,"", vXMLRequestEnvelope) 'type value in element xmltem:Cd_proposta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtem_succod", vtem_succod,"", vXMLRequestEnvelope) 'type value in element xmltem:succod
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtem_ramcod", vtem_ramcod,"", vXMLRequestEnvelope) 'type value in element xmltem:ramcod
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtem_aplnumdig", vtem_aplnumdig,"", vXMLRequestEnvelope) 'type value in element xmltem:aplnumdig
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtem_prporg", vtem_prporg,"", vXMLRequestEnvelope) 'type value in element xmltem:prporg
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtem_prpnumdig", vtem_prpnumdig,"", vXMLRequestEnvelope) 'type value in element xmltem:prpnumdig
							Test.TestLog("Consumir o serviço ConsultaProposta", "O serviço ConsultaProposta deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço ConsultaProposta foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService OSB LocalizaPropESEG
                            If p_xmlSoapLeanTest.StatusCode = 200 Then
                                Test.TestLog("Verificar o retorno do serviço ConsultaProposta", "O serviço ConsultaProposta retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço ConsultaProposta foi consumido com sucesso", typelog.Passed, , False)
                            Else
                                Test.TestLog("Verificar o retorno do serviço ConsultaProposta", "O serviço ConsultaProposta retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço ConsultaProposta não foi consumido com sucesso", typelog.Failed, , False)
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
				HandlerError("test_OSB_LocalizaPropESEG_ConsultaProposta.test_OSB_LocalizaPropESEG_ConsultaProposta.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vtem_Cd_filial, vtem_Cd_produto, vtem_Cd_proposta, vtem_succod, vtem_ramcod, vtem_aplnumdig, vtem_prporg, vtem_prpnumdig, vIsOpenSystem As String

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
					vtem_Cd_filial = pc_db.Fieldt("vtem_Cd_filial")
					vtem_Cd_produto = pc_db.Fieldt("vtem_Cd_produto")
					vtem_Cd_proposta = pc_db.Fieldt("vtem_Cd_proposta")
					vtem_succod = pc_db.Fieldt("vtem_succod")
					vtem_ramcod = pc_db.Fieldt("vtem_ramcod")
					vtem_aplnumdig = pc_db.Fieldt("vtem_aplnumdig")
					vtem_prporg = pc_db.Fieldt("vtem_prporg")
					vtem_prpnumdig = pc_db.Fieldt("vtem_prpnumdig")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_OSB_LocalizaPropESEG_ConsultaProposta.test_OSB_LocalizaPropESEG_ConsultaProposta.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
