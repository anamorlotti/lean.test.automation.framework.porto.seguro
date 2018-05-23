'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 21/03/2018 16:29:45 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... Normas OSB
'Functionality:... validarNormasRecebimento
'Master Test:..... No Defined
'TableTest:....... test_Normas_OSB_validarNormasRecebimento
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Normas_OSB_validarNormasRecebimento
    Public Class test_Normas_OSB_validarNormasRecebimento
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "https://osbhmlcorp.portoseguro.brasil:443/NormasIntegrationService/NormasIntegrationServiceSoap11V1_0"'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters Normas OSB
							vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope Normas OSB
							vXMLRequestEnvelope = Test.SetValueXMLElement("vweb_numeroCgcCpf", vweb_numeroCgcCpf,"", vXMLRequestEnvelope) 'type value in xml elementweb:numeroCgcCpf
							vXMLRequestEnvelope = Test.SetValueXMLElement("vweb_ordemCgcCpf", vweb_ordemCgcCpf,"", vXMLRequestEnvelope) 'type value in xml elementweb:ordemCgcCpf
							vXMLRequestEnvelope = Test.SetValueXMLElement("vweb_digitoCgcCpf", vweb_digitoCgcCpf,"", vXMLRequestEnvelope) 'type value in xml elementweb:digitoCgcCpf
							vXMLRequestEnvelope = Test.SetValueXMLElement("vweb_codigoBanco", vweb_codigoBanco,"", vXMLRequestEnvelope) 'type value in xml elementweb:codigoBanco
							vXMLRequestEnvelope = Test.SetValueXMLElement("vweb_numeroAgencia", vweb_numeroAgencia,"", vXMLRequestEnvelope) 'type value in xml elementweb:numeroAgencia
							vXMLRequestEnvelope = Test.SetValueXMLElement("vweb_numeroConta", vweb_numeroConta,"", vXMLRequestEnvelope) 'type value in xml elementweb:numeroConta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vweb_codigoSistema", vweb_codigoSistema,"", vXMLRequestEnvelope) 'type value in xml elementweb:codigoSistema
							vXMLRequestEnvelope = Test.SetValueXMLElement("vweb_tipoDocumento", vweb_tipoDocumento,"", vXMLRequestEnvelope) 'type value in xml elementweb:tipoDocumento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vweb_codigoOrigem", vweb_codigoOrigem,"", vXMLRequestEnvelope) 'type value in xml elementweb:codigoOrigem
							vXMLRequestEnvelope = Test.SetValueXMLElement("vweb_numeroProposta", vweb_numeroProposta,"", vXMLRequestEnvelope) 'type value in xml elementweb:numeroProposta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vweb_numeroDocumento", vweb_numeroDocumento,"", vXMLRequestEnvelope) 'type value in xml elementweb:numero Documento
							
							Test.TestLog("Consumir o serviço validarNormasRecebimento", "O serviço validarNormasRecebimento deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço validarNormasRecebimento foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService Normas OSB
                            If p_xmlSoapLeanTest.StatusCode = 200 Then
                                Test.TestLog("Verificar o retorno do serviço validarNormasRecebimento", "O serviço validarNormasRecebimento retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço validarNormasRecebimento foi consumido com sucesso", typelog.Passed, , False)
                            Else
                                Test.TestLog("Verificar o retorno do serviço validarNormasRecebimento", "O serviço validarNormasRecebimento retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço validarNormasRecebimento não foi consumido com sucesso", typelog.Failed, , False)
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
				HandlerError("test_Normas_OSB_validarNormasRecebimento.test_Normas_OSB_validarNormasRecebimento.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vweb_numeroDocumento, vXMLParameters,vXMLRequestEnvelope,vweb_numeroCgcCpf, vweb_ordemCgcCpf, vweb_digitoCgcCpf, vweb_codigoBanco, vweb_numeroAgencia, vweb_numeroConta, vweb_codigoSistema, vweb_tipoDocumento, vweb_codigoOrigem, vweb_numeroProposta, vIsOpenSystem As String

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
					vweb_numeroCgcCpf = pc_db.Fieldt("vweb_numeroCgcCpf")
					vweb_ordemCgcCpf = pc_db.Fieldt("vweb_ordemCgcCpf")
					vweb_digitoCgcCpf = pc_db.Fieldt("vweb_digitoCgcCpf")
					vweb_codigoBanco = pc_db.Fieldt("vweb_codigoBanco")
					vweb_numeroAgencia = pc_db.Fieldt("vweb_numeroAgencia")
					vweb_numeroConta = pc_db.Fieldt("vweb_numeroConta")
					vweb_codigoSistema = pc_db.Fieldt("vweb_codigoSistema")
					vweb_tipoDocumento = pc_db.Fieldt("vweb_tipoDocumento")
					vweb_codigoOrigem = pc_db.Fieldt("vweb_codigoOrigem")
					vweb_numeroProposta = pc_db.Fieldt("vweb_numeroProposta")
					vweb_numeroDocumento = pc_db.Fieldt("vweb_numeroDocumento")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Normas_OSB_validarNormasRecebimento.test_Normas_OSB_validarNormasRecebimento.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
