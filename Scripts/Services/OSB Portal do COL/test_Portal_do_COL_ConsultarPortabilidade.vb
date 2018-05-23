'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 20/03/2018 11:49:59 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... Portal do COL
'Functionality:... ConsultarPortabilidade
'Master Test:..... No Defined
'TableTest:....... test_Portal_do_COL_ConsultarPortabilidade
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Portal_do_COL_ConsultarPortabilidade
    Public Class test_Portal_do_COL_ConsultarPortabilidade
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "http://osbhmlcrp.portoseguro.brasil:80/PortalCorretorIntegrationService/Proxy_Services/PrevidenciaPortalCorretorServiceSoap11V1_0"'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters Portal do COL
							vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope Portal do COL
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_susepCorretor", vv11_susepCorretor,"", vXMLRequestEnvelope) 'type value in xml elementv11:susepCorretor
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_cpfParticipante", vv11_cpfParticipante,"", vXMLRequestEnvelope) 'type value in xml elementv11:cpfParticipante
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_numeroPortabilidade", vv11_numeroPortabilidade,"", vXMLRequestEnvelope) 'type value in xml elementv11:numeroPortabilidade
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_codigoOrigem", vv11_codigoOrigem,"", vXMLRequestEnvelope) 'type value in xml elementv11:codigoOrigem
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_numeroProposta", vv11_numeroProposta,"", vXMLRequestEnvelope) 'type value in xml elementv11:numeroProposta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_nomeParticipante", vv11_nomeParticipante,"", vXMLRequestEnvelope) 'type value in xml elementv11:nomeParticipante
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_dataInicial", vv11_dataInicial,"", vXMLRequestEnvelope) 'type value in xml elementv11:dataInicial
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_dataFinal", vv11_dataFinal,"", vXMLRequestEnvelope) 'type value in xml elementv11:dataFinal
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_tipo", vv11_tipo,"", vXMLRequestEnvelope) 'type value in xml elementv11:tipo
							Test.TestLog("Consumir o serviço ConsultarPortabilidade", "O serviço ConsultarPortabilidade deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço ConsultarPortabilidade foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService Portal do COL
                            If p_xmlSoapLeanTest.StatusCode = 200 Then
                                Test.TestLog("Verificar o retorno do serviço ConsultarPortabilidade", "O serviço ConsultarPortabilidade retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço ConsultarPortabilidade foi consumido com sucesso", typelog.Passed, , False)
                            Else
                                Test.TestLog("Verificar o retorno do serviço ConsultarPortabilidade", "O serviço ConsultarPortabilidade retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço ConsultarPortabilidade não foi consumido com sucesso", typelog.Failed, , False)
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
				HandlerError("test_Portal_do_COL_ConsultarPortabilidade.test_Portal_do_COL_ConsultarPortabilidade.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vv11_susepCorretor, vv11_cpfParticipante, vv11_numeroPortabilidade, vv11_codigoOrigem, vv11_numeroProposta, vv11_nomeParticipante, vv11_dataInicial, vv11_dataFinal, vv11_tipo, vIsOpenSystem As String

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
					vv11_susepCorretor = pc_db.Fieldt("vv11_susepCorretor")
					vv11_cpfParticipante = pc_db.Fieldt("vv11_cpfParticipante")
					vv11_numeroPortabilidade = pc_db.Fieldt("vv11_numeroPortabilidade")
					vv11_codigoOrigem = pc_db.Fieldt("vv11_codigoOrigem")
					vv11_numeroProposta = pc_db.Fieldt("vv11_numeroProposta")
					vv11_nomeParticipante = pc_db.Fieldt("vv11_nomeParticipante")
					vv11_dataInicial = pc_db.Fieldt("vv11_dataInicial")
					vv11_dataFinal = pc_db.Fieldt("vv11_dataFinal")
					vv11_tipo = pc_db.Fieldt("vv11_tipo")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Portal_do_COL_ConsultarPortabilidade.test_Portal_do_COL_ConsultarPortabilidade.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
