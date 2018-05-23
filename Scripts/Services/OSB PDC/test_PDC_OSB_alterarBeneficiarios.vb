'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 22/03/2018 10:12:31 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... PDC OSB
'Functionality:... alterarBeneficiarios
'Master Test:..... No Defined
'TableTest:....... test_PDC_OSB_alterarBeneficiarios
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_PDC_OSB_alterarBeneficiarios
    Public Class test_PDC_OSB_alterarBeneficiarios
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
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodSucursal", vcodSucursal,"", vXMLRequestEnvelope) 'type value in xml elementcodSucursal
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodRamo", vcodRamo,"", vXMLRequestEnvelope) 'type value in xml elementcodRamo
							vXMLRequestEnvelope = Test.SetValueXMLElement("vseqApolice", vseqApolice,"", vXMLRequestEnvelope) 'type value in xml elementseqApolice
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodOrigem", vcodOrigem,"", vXMLRequestEnvelope) 'type value in xml elementcodOrigem
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodProposta", vcodProposta,"", vXMLRequestEnvelope) 'type value in xml elementcodProposta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodBeneficiario", vcodBeneficiario,"", vXMLRequestEnvelope) 'type value in xml elementcodBeneficiario
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnomeBeneficiario", vnomeBeneficiario,"", vXMLRequestEnvelope) 'type value in xml elementnomeBeneficiario
							Test.TestLog("Consumir o serviço alterarBeneficiarios", "O serviço alterarBeneficiarios deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço alterarBeneficiarios foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService PDC OSB
                            If p_xmlSoapLeanTest.StatusCode = 200 Then
                                Test.TestLog("Verificar o retorno do serviço alterarBeneficiarios", "O serviço alterarBeneficiarios retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço alterarBeneficiarios foi consumido com sucesso", typelog.Passed, , False)
                            Else
                                Test.TestLog("Verificar o retorno do serviço alterarBeneficiarios", "O serviço alterarBeneficiarios retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço alterarBeneficiarios não foi consumido com sucesso", typelog.Failed, , False)
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
				HandlerError("test_PDC_OSB_alterarBeneficiarios.test_PDC_OSB_alterarBeneficiarios.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vcodSucursal, vcodRamo, vseqApolice, vcodOrigem, vcodProposta, vcodBeneficiario, vnomeBeneficiario, vIsOpenSystem As String

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
					vcodBeneficiario = pc_db.Fieldt("vcodBeneficiario")
					vnomeBeneficiario = pc_db.Fieldt("vnomeBeneficiario")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_PDC_OSB_alterarBeneficiarios.test_PDC_OSB_alterarBeneficiarios.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
