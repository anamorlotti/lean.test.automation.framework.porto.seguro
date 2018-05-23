'**********************************************************************************************************************************************************************
'Create by LeanTest Automation 3.8 in 10/04/2018 18:09:48 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... PDC OSB
'Functionality:... alterarDadosParticipante
'Master Test:..... No Defined
'TableTest:....... test_PDC_OSB_alterarDadosParticipante
'**********************************************************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_PDC_OSB_alterarDadosParticipante
    Public Class test_PDC_OSB_alterarDadosParticipante
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "http://osbhmlcrp.portoseguro.brasil:80/PortalClienteIntegrationService/Proxy_Services/PrevidenciaPortalClienteServiceV1_0"'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters PDC OSB
							vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope PDC OSB
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodSucursal", vcodSucursal,"", vXMLRequestEnvelope) 'type value in element xmlcodSucursal
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodRamo", vcodRamo,"", vXMLRequestEnvelope) 'type value in element xmlcodRamo
							vXMLRequestEnvelope = Test.SetValueXMLElement("vseqApolice", vseqApolice,"", vXMLRequestEnvelope) 'type value in element xmlseqApolice
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodOrigem", vcodOrigem,"", vXMLRequestEnvelope) 'type value in element xmlcodOrigem
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodProposta", vcodProposta,"", vXMLRequestEnvelope) 'type value in element xmlcodProposta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnomeTitular", vnomeTitular,"", vXMLRequestEnvelope) 'type value in element xmlnomeTitular
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodCep", vcodCep,"", vXMLRequestEnvelope) 'type value in element xmlcodCep
							vXMLRequestEnvelope = Test.SetValueXMLElement("vdescLogradouro", vdescLogradouro,"", vXMLRequestEnvelope) 'type value in element xmldescLogradouro
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroLogradouro", vnumeroLogradouro,"", vXMLRequestEnvelope) 'type value in element xmlnumeroLogradouro
							vXMLRequestEnvelope = Test.SetValueXMLElement("vdescComplemento", vdescComplemento,"", vXMLRequestEnvelope) 'type value in element xmldescComplemento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vdescBairro", vdescBairro,"", vXMLRequestEnvelope) 'type value in element xmldescBairro
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodDddResi", vcodDddResi,"", vXMLRequestEnvelope) 'type value in element xmlcodDddResi
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroTelefoneResi", vnumeroTelefoneResi,"", vXMLRequestEnvelope) 'type value in element xmlnumeroTelefoneResi
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodDddCome", vcodDddCome,"", vXMLRequestEnvelope) 'type value in element xmlcodDddCome
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroTelefoneCome", vnumeroTelefoneCome,"", vXMLRequestEnvelope) 'type value in element xmlnumeroTelefoneCome
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodDddCelu", vcodDddCelu,"", vXMLRequestEnvelope) 'type value in element xmlcodDddCelu
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroTelefoneCelu", vnumeroTelefoneCelu,"", vXMLRequestEnvelope) 'type value in element xmlnumeroTelefoneCelu
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodEstadoCivil", vcodEstadoCivil,"", vXMLRequestEnvelope) 'type value in element xmlcodEstadoCivil
							vXMLRequestEnvelope = Test.SetValueXMLElement("vdescEmail", vdescEmail,"", vXMLRequestEnvelope) 'type value in element xmldescEmail
							Test.TestLog("Consumir o serviço alterarDadosParticipante", "O serviço alterarDadosParticipante deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço alterarDadosParticipante foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService PDC OSB
							If p_xmlSoapLeanTest.StatusCode = 200 Then
								Test.TestLog("Verificar o retorno do serviço alterarDadosParticipante", "O serviço alterarDadosParticipante retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço alterarDadosParticipante foi consumido com sucesso", typelog.Passed, , False)
							Else
							Test.TestLog("Verificar o retorno do serviço alterarDadosParticipante", "O serviço alterarDadosParticipante retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço alterarDadosParticipante não foi consumido com sucesso", typelog.Failed, , False)
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
		HandlerError("test_PDC_OSB_alterarDadosParticipante.test_PDC_OSB_alterarDadosParticipante.Run: " & ex.Message)
		Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
		Return False
            End Try
        End Function

        '****************************************************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vcodSucursal, vcodRamo, vseqApolice, vcodOrigem, vcodProposta, vnomeTitular, vcodCep, vdescLogradouro, vnumeroLogradouro, vdescComplemento, vdescBairro, vcodDddResi, vnumeroTelefoneResi, vcodDddCome, vnumeroTelefoneCome, vcodDddCelu, vnumeroTelefoneCelu, vcodEstadoCivil, vdescEmail, vIsOpenSystem As String

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
					vcodSucursal = pc_db.Fieldt("vcodSucursal")
					vcodRamo = pc_db.Fieldt("vcodRamo")
					vseqApolice = pc_db.Fieldt("vseqApolice")
					vcodOrigem = pc_db.Fieldt("vcodOrigem")
					vcodProposta = pc_db.Fieldt("vcodProposta")
					vnomeTitular = pc_db.Fieldt("vnomeTitular")
					vcodCep = pc_db.Fieldt("vcodCep")
					vdescLogradouro = pc_db.Fieldt("vdescLogradouro")
					vnumeroLogradouro = pc_db.Fieldt("vnumeroLogradouro")
					vdescComplemento = pc_db.Fieldt("vdescComplemento")
					vdescBairro = pc_db.Fieldt("vdescBairro")
					vcodDddResi = pc_db.Fieldt("vcodDddResi")
					vnumeroTelefoneResi = pc_db.Fieldt("vnumeroTelefoneResi")
					vcodDddCome = pc_db.Fieldt("vcodDddCome")
					vnumeroTelefoneCome = pc_db.Fieldt("vnumeroTelefoneCome")
					vcodDddCelu = pc_db.Fieldt("vcodDddCelu")
					vnumeroTelefoneCelu = pc_db.Fieldt("vnumeroTelefoneCelu")
					vcodEstadoCivil = pc_db.Fieldt("vcodEstadoCivil")
					vdescEmail = pc_db.Fieldt("vdescEmail")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_PDC_OSB_alterarDadosParticipante.test_PDC_OSB_alterarDadosParticipante.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
