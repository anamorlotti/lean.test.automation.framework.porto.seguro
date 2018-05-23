'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 20/03/2018 11:37:30 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... Portal do COL
'Functionality:... ConsultarPropostas
'Master Test:..... No Defined
'TableTest:....... test_Portal_do_COL_ConsultarPropostas
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Portal_do_COL_ConsultarPropostas
    Public Class test_Portal_do_COL_ConsultarPropostas
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
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_susepCorretor", vv1_susepCorretor,"", vXMLRequestEnvelope) 'type value in xml elementv1:susepCorretor
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_dataInicial", vv1_dataInicial,"", vXMLRequestEnvelope) 'type value in xml elementv1:dataInicial
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_dataFinal", vv1_dataFinal,"", vXMLRequestEnvelope) 'type value in xml elementv1:dataFinal
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_status", vv1_status,"", vXMLRequestEnvelope) 'type value in xml elementv1:status
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_nomeParticipante", vv1_nomeParticipante,"", vXMLRequestEnvelope) 'type value in xml elementv1:nomeParticipante
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_codigoPlano", vv1_codigoPlano,"", vXMLRequestEnvelope) 'type value in xml elementv1:codigoPlano
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_codigoOrigem", vv1_codigoOrigem,"", vXMLRequestEnvelope) 'type value in xml elementv1:codigoOrigem
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_numeroProposta", vv1_numeroProposta,"", vXMLRequestEnvelope) 'type value in xml elementv1:numeroProposta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_cpfParticipante", vv1_cpfParticipante,"", vXMLRequestEnvelope) 'type value in xml elementv1:cpfParticipante
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_considerarSomenteMovtosSaidaTotal", vv1_considerarSomenteMovtosSaidaTotal,"", vXMLRequestEnvelope) 'type value in xml elementv1:considerarSomenteMovtosSaidaTotal
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_codigoSucursal", vv1_codigoSucursal,"", vXMLRequestEnvelope) 'type value in xml elementv1:codigoSucursal
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_codigoRamo", vv1_codigoRamo,"", vXMLRequestEnvelope) 'type value in xml elementv1:codigoRamo
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_codigoApolice", vv1_codigoApolice,"", vXMLRequestEnvelope) 'type value in xml elementv1:codigoApolice
							Test.TestLog("Consumir o serviço ConsultarPropostas", "O serviço ConsultarPropostas deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço ConsultarPropostas foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService Portal do COL
                            If p_xmlSoapLeanTest.StatusCode = 200 Then
                                Test.TestLog("Verificar o retorno do serviço ConsultarPropostas", "O serviço ConsultarPropostas retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço ConsultarPropostas foi consumido com sucesso", typelog.Passed, , False)
                            Else
                                Test.TestLog("Verificar o retorno do serviço ConsultarPropostas", "O serviço ConsultarPropostas retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço ConsultarPropostas não foi consumido com sucesso", typelog.Failed, , False)
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
				HandlerError("test_Portal_do_COL_ConsultarPropostas.test_Portal_do_COL_ConsultarPropostas.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vv1_susepCorretor, vv1_dataInicial, vv1_dataFinal, vv1_status, vv1_nomeParticipante, vv1_codigoPlano, vv1_codigoOrigem, vv1_numeroProposta, vv1_cpfParticipante, vv1_considerarSomenteMovtosSaidaTotal, vv1_codigoSucursal, vv1_codigoRamo, vv1_codigoApolice, vIsOpenSystem As String

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
					vv1_susepCorretor = pc_db.Fieldt("vv1_susepCorretor")
					vv1_dataInicial = pc_db.Fieldt("vv1_dataInicial")
					vv1_dataFinal = pc_db.Fieldt("vv1_dataFinal")
					vv1_status = pc_db.Fieldt("vv1_status")
					vv1_nomeParticipante = pc_db.Fieldt("vv1_nomeParticipante")
					vv1_codigoPlano = pc_db.Fieldt("vv1_codigoPlano")
					vv1_codigoOrigem = pc_db.Fieldt("vv1_codigoOrigem")
					vv1_numeroProposta = pc_db.Fieldt("vv1_numeroProposta")
					vv1_cpfParticipante = pc_db.Fieldt("vv1_cpfParticipante")
					vv1_considerarSomenteMovtosSaidaTotal = pc_db.Fieldt("vv1_considerarSomenteMovtosSaidaTotal")
					vv1_codigoSucursal = pc_db.Fieldt("vv1_codigoSucursal")
					vv1_codigoRamo = pc_db.Fieldt("vv1_codigoRamo")
					vv1_codigoApolice = pc_db.Fieldt("vv1_codigoApolice")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Portal_do_COL_ConsultarPropostas.test_Portal_do_COL_ConsultarPropostas.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
