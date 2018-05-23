'**********************************************************************************************************************************************************************
'Create by LeanTest Automation 3.8 in 23/03/2018 14:59:41 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... Serasa
'Functionality:... obterDadosSerasa
'Master Test:..... No Defined
'TableTest:....... test_Serasa_obterDadosSerasa
'**********************************************************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Serasa_obterDadosSerasa
    Public Class test_Serasa_obterDadosSerasa
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "https://osbhmlcorp.portoseguro.brasil:443/SerasaIntegrationService/Proxy_Services/SerasaIntegrationServiceSoap11V3_0"'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters Serasa
							vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope Serasa
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_cnpjCpf", vv11_cnpjCpf,"", vXMLRequestEnvelope) 'type value in xml elementv11:cnpjCpf
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_tipoPessoa", vv11_tipoPessoa,"", vXMLRequestEnvelope) 'type value in xml elementv11:tipoPessoa
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_siglaDepartamento", vv11_siglaDepartamento,"", vXMLRequestEnvelope) 'type value in xml elementv11:siglaDepartamento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_codigoRamo", vv11_codigoRamo,"", vXMLRequestEnvelope) 'type value in xml elementv11:codigoRamo
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_idadeInformacao", vv11_idadeInformacao,"", vXMLRequestEnvelope) 'type value in xml elementv11:idadeInformacao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_tipoSistema", vv11_tipoSistema,"", vXMLRequestEnvelope) 'type value in xml elementv11:tipoSistema
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_nomePessoa", vv11_nomePessoa,"", vXMLRequestEnvelope) 'type value in xml elementv11:nomePessoa
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_sistemaOrigem", vv11_sistemaOrigem,"", vXMLRequestEnvelope) 'type value in xml elementv11:sistemaOrigem
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_codigoProduto", vv11_codigoProduto,"", vXMLRequestEnvelope) 'type value in xml elementv11:codigoProduto
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_xmlMQ", vv11_xmlMQ,"", vXMLRequestEnvelope) 'type value in xml elementv11:xmlMQ
							Test.TestLog("Consumir o serviço obterDadosSerasa", "O serviço obterDadosSerasa deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço obterDadosSerasa foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService Serasa
                            If p_xmlSoapLeanTest.StatusCode = 200 Then
                                Test.TestLog("Verificar o retorno do serviço obterDadosSerasa", "O serviço obterDadosSerasa retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço obterDadosSerasa foi consumido com sucesso", typelog.Passed, , False)
                            Else
                                Test.TestLog("Verificar o retorno do serviço obterDadosSerasa", "O serviço obterDadosSerasa retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço obterDadosSerasa não foi consumido com sucesso", typelog.Failed, , False)
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
		HandlerError("test_Serasa_obterDadosSerasa.test_Serasa_obterDadosSerasa.Run: " & ex.Message)
		Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
		Return False
            End Try
        End Function

        '****************************************************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vv11_cnpjCpf, vv11_tipoPessoa, vv11_siglaDepartamento, vv11_codigoRamo, vv11_idadeInformacao, vv11_tipoSistema, vv11_nomePessoa, vv11_sistemaOrigem, vv11_codigoProduto, vv11_xmlMQ, vIsOpenSystem As String

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
					vv11_cnpjCpf = pc_db.Fieldt("vv11_cnpjCpf")
					vv11_tipoPessoa = pc_db.Fieldt("vv11_tipoPessoa")
					vv11_siglaDepartamento = pc_db.Fieldt("vv11_siglaDepartamento")
					vv11_codigoRamo = pc_db.Fieldt("vv11_codigoRamo")
					vv11_idadeInformacao = pc_db.Fieldt("vv11_idadeInformacao")
					vv11_tipoSistema = pc_db.Fieldt("vv11_tipoSistema")
					vv11_nomePessoa = pc_db.Fieldt("vv11_nomePessoa")
					vv11_sistemaOrigem = pc_db.Fieldt("vv11_sistemaOrigem")
					vv11_codigoProduto = pc_db.Fieldt("vv11_codigoProduto")
					vv11_xmlMQ = pc_db.Fieldt("vv11_xmlMQ")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Serasa_obterDadosSerasa.test_Serasa_obterDadosSerasa.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
