'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 23/03/2018 11:36:47 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... OSB Envio Email
'Functionality:... enviarEmail
'Master Test:..... No Defined
'TableTest:....... test_OSB_Envio_Email_enviarEmail
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_OSB_Envio_Email_enviarEmail
    Public Class test_OSB_Envio_Email_enviarEmail
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "https://osbhmlcorp.portoseguro.brasil/EmailCommonService/Proxy_Services/ProxyEmailServiceSOAPV1_0?wsdl"'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters OSB Envio Email
							vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope OSB Envio Email
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmail_idAplicacaoUtilizadora", vmail_idAplicacaoUtilizadora,"", vXMLRequestEnvelope) 'type value in xml elementmail:idAplicacaoUtilizadora
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmail_unidadeNegocio", vmail_unidadeNegocio,"", vXMLRequestEnvelope) 'type value in xml elementmail:unidadeNegocio
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmail_centroDeCustoPagador", vmail_centroDeCustoPagador,"", vXMLRequestEnvelope) 'type value in xml elementmail:centroDeCustoPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmail_emailDoResponsavelEmEnviar", vmail_emailDoResponsavelEmEnviar,"", vXMLRequestEnvelope) 'type value in xml elementmail:emailDoResponsavelEmEnviar
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmail_emailDoSegundoResponsavelEmEnviar", vmail_emailDoSegundoResponsavelEmEnviar,"", vXMLRequestEnvelope) 'type value in xml elementmail:emailDoSegundoResponsavelEmEnviar
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmail_emailRecebedor", vmail_emailRecebedor,"", vXMLRequestEnvelope) 'type value in xml elementmail:emailRecebedor
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmail_emailResponsavelPorResponder", vmail_emailResponsavelPorResponder,"", vXMLRequestEnvelope) 'type value in xml elementmail:emailResponsavelPorResponder
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmail_assuntoEmail", vmail_assuntoEmail,"", vXMLRequestEnvelope) 'type value in xml elementmail:assuntoEmail
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmail_tipoCorpoEmail", vmail_tipoCorpoEmail,"", vXMLRequestEnvelope) 'type value in xml elementmail:tipoCorpoEmail
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmail_corpoEmail", vmail_corpoEmail,"", vXMLRequestEnvelope) 'type value in xml elementmail:corpoEmail
							Test.TestLog("Consumir o serviço enviarEmail", "O serviço enviarEmail deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço enviarEmail foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService OSB Envio Email
                            If p_xmlSoapLeanTest.StatusCode = 200 Then
                                Test.TestLog("Verificar o retorno do serviço enviarEmail", "O serviço enviarEmail retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço enviarEmail foi consumido com sucesso", typelog.Passed, , False)
                            Else
                                Test.TestLog("Verificar o retorno do serviço enviarEmail", "O serviço enviarEmail retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço enviarEmail não foi consumido com sucesso", typelog.Failed, , False)
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
				HandlerError("test_OSB_Envio_Email_enviarEmail.test_OSB_Envio_Email_enviarEmail.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vmail_idAplicacaoUtilizadora, vmail_unidadeNegocio, vmail_centroDeCustoPagador, vmail_emailDoResponsavelEmEnviar, vmail_emailDoSegundoResponsavelEmEnviar, vmail_emailRecebedor, vmail_emailResponsavelPorResponder, vmail_assuntoEmail, vmail_tipoCorpoEmail, vmail_corpoEmail, vIsOpenSystem As String

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
					vmail_idAplicacaoUtilizadora = pc_db.Fieldt("vmail_idAplicacaoUtilizadora")
					vmail_unidadeNegocio = pc_db.Fieldt("vmail_unidadeNegocio")
					vmail_centroDeCustoPagador = pc_db.Fieldt("vmail_centroDeCustoPagador")
					vmail_emailDoResponsavelEmEnviar = pc_db.Fieldt("vmail_emailDoResponsavelEmEnviar")
					vmail_emailDoSegundoResponsavelEmEnviar = pc_db.Fieldt("vmail_emailDoSegundoResponsavelEmEnviar")
					vmail_emailRecebedor = pc_db.Fieldt("vmail_emailRecebedor")
					vmail_emailResponsavelPorResponder = pc_db.Fieldt("vmail_emailResponsavelPorResponder")
					vmail_assuntoEmail = pc_db.Fieldt("vmail_assuntoEmail")
					vmail_tipoCorpoEmail = pc_db.Fieldt("vmail_tipoCorpoEmail")
					vmail_corpoEmail = pc_db.Fieldt("vmail_corpoEmail")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_OSB_Envio_Email_enviarEmail.test_OSB_Envio_Email_enviarEmail.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
