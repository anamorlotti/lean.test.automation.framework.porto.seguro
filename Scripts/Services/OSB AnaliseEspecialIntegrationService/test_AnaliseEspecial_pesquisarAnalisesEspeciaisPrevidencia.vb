'**********************************************************************************************************************************************************************
'Create by LeanTest Automation 3.8 in 27/03/2018 15:05:12 (By LeanTest Automation Test) 
'**********************************************************************************************************************************************************************
'Create by LeanTest Automation 3.8 in 03/04/2018 18:20:06 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... AnaliseEspecial
'Functionality:... pesquisarAnalisesEspeciaisPrevidencia
'Master Test:..... No Defined
'TableTest:....... test_AnaliseEspecial_pesquisarAnalisesEspeciaisPrevidencia
'**********************************************************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_AnaliseEspecial_pesquisarAnalisesEspeciaisPrevidencia
    Public Class test_AnaliseEspecial_pesquisarAnalisesEspeciaisPrevidencia
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
                                p_pathUrlApp = "https://osbtstcorp.portoseguro.brasil:443/AnaliseEspecialIntegrationService/Proxy_Services/AnaliseEspecialIntegrationServiceSOAP11V1_0" 'xml.Read("urlLocal") 'Create urlLocal element in XML
                                If Not Test.StartTests Then Return False
                            End If
                            vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters AnaliseEspecial
                            vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope AnaliseEspecial
                            vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_cpf", vv1_cpf, "", vXMLRequestEnvelope) 'type value in element xmlv1:cpf
                            vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_digito", vv1_digito, "", vXMLRequestEnvelope) 'type value in element xmlv1:digito
                            vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_cgcord", vv1_cgcord, "", vXMLRequestEnvelope) 'type value in element xmlv1:cgcord
                            vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_origemPropostaPrincipal", vv1_origemPropostaPrincipal, "", vXMLRequestEnvelope) 'type value in element xmlv1:origemPropostaPrincipal
                            vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_numeroPropostaPrincipal", vv1_numeroPropostaPrincipal, "", vXMLRequestEnvelope) 'type value in element xmlv1:numeroPropostaPrincipal
                            vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_dataNascimento", vv1_dataNascimento, "", vXMLRequestEnvelope) 'type value in element xmlv1:dataNascimento
                            vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_statusAnalise", vv1_statusAnalise, "", vXMLRequestEnvelope) 'type value in element xmlv1:statusAnalise
                            vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_rg", vv1_rg, "", vXMLRequestEnvelope) 'type value in element xmlv1:rg
                            vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_ramo", vv1_ramo, "", vXMLRequestEnvelope) 'type value in element xmlv1:ramo
                            vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_codigoUnidadeOperacional", vv1_codigoUnidadeOperacional, "", vXMLRequestEnvelope) 'type value in element xmlv1:codigoUnidadeOperacional
                            vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_nomeSegurado", vv1_nomeSegurado, "", vXMLRequestEnvelope) 'type value in element xmlv1:nomeSegurado
                            vXMLRequestEnvelope = Test.SetValueXMLElement("vv1_valorImportanciaSegurada", vv1_valorImportanciaSegurada, "", vXMLRequestEnvelope) 'type value in element xmlv1:valorImportanciaSegurada
                            Test.TestLog("Consumir o serviço pesquisarAnalisesEspeciaisPrevidencia", "O serviço pesquisarAnalisesEspeciaisPrevidencia deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço pesquisarAnalisesEspeciaisPrevidencia foi consumido com sucesso", typelog.Passed, , False)
                            p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService AnaliseEspecial
                            If p_xmlSoapLeanTest.StatusCode = 200 Then
                                Test.TestLog("Verificar o retorno do serviço pesquisarAnalisesEspeciaisPrevidencia", "O serviço pesquisarAnalisesEspeciaisPrevidencia retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço pesquisarAnalisesEspeciaisPrevidencia foi consumido com sucesso", typelog.Passed, , False)
                            Else
                                Test.TestLog("Verificar o retorno do serviço pesquisarAnalisesEspeciaisPrevidencia", "O serviço pesquisarAnalisesEspeciaisPrevidencia retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço pesquisarAnalisesEspeciaisPrevidencia não foi consumido com sucesso", typelog.Failed, , False)
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
                HandlerError("test_AnaliseEspecial_pesquisarAnalisesEspeciaisPrevidencia.test_AnaliseEspecial_pesquisarAnalisesEspeciaisPrevidencia.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '****************************************************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters, vXMLRequestEnvelope, vv1_cpf, vv1_digito, vv1_cgcord, vv1_origemPropostaPrincipal, vv1_numeroPropostaPrincipal, vv1_dataNascimento, vv1_statusAnalise, vv1_rg, vv1_ramo, vv1_codigoUnidadeOperacional, vv1_nomeSegurado, vv1_valorImportanciaSegurada, vIsOpenSystem As String

        Private Function StartTest() As Boolean
            Dim strQueryOut1, strQueryOut2, strQueryOut3, strQueryOut4, strQueryOut5, strQueryOut6 As String
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
                    vv1_cpf = pc_db.Fieldt("vv1_cpf")
                    vv1_digito = pc_db.Fieldt("vv1_digito")
                    vv1_cgcord = pc_db.Fieldt("vv1_cgcord")
                    vv1_origemPropostaPrincipal = pc_db.Fieldt("vv1_origemPropostaPrincipal")
                    vv1_numeroPropostaPrincipal = pc_db.Fieldt("vv1_numeroPropostaPrincipal")
                    vv1_dataNascimento = pc_db.Fieldt("vv1_dataNascimento")
                    vv1_statusAnalise = pc_db.Fieldt("vv1_statusAnalise")
                    vv1_rg = pc_db.Fieldt("vv1_rg")
                    vv1_ramo = pc_db.Fieldt("vv1_ramo")
                    vv1_codigoUnidadeOperacional = pc_db.Fieldt("vv1_codigoUnidadeOperacional")
                    vv1_nomeSegurado = pc_db.Fieldt("vv1_nomeSegurado")
                    vv1_valorImportanciaSegurada = pc_db.Fieldt("vv1_valorImportanciaSegurada")
                    vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")


                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_AnaliseEspecial_pesquisarAnalisesEspeciaisPrevidencia.test_AnaliseEspecial_pesquisarAnalisesEspeciaisPrevidencia.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
