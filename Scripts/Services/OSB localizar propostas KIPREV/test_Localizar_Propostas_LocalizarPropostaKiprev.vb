'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 19/04/2018 10:49:15 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... Localizar Propostas
'Functionality:... LocalizarPropostaKiprev
'Master Test:..... No Defined
'TableTest:....... test_Localizar_Propostas_LocalizarPropostaKiprev
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Localizar_Propostas_LocalizarPropostaKiprev
    Public Class test_Localizar_Propostas_LocalizarPropostaKiprev
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "http://nt2621.portoseguro.brasil:7002/ki-ws/LocalizadorKiprevWS?wsdl"'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters Localizar Propostas
							vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope Localizar Propostas
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcor_Cd_filial", vcor_Cd_filial,"", vXMLRequestEnvelope) 'type value in element xmlcor:Cd_filial
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcor_Cd_produto", vcor_Cd_produto,"", vXMLRequestEnvelope) 'type value in element xmlcor:Cd_produto
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcor_Cd_proposta", vcor_Cd_proposta,"", vXMLRequestEnvelope) 'type value in element xmlcor:Cd_proposta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcor_succod", vcor_succod,"", vXMLRequestEnvelope) 'type value in element xmlcor:succod
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcor_ramcod", vcor_ramcod,"", vXMLRequestEnvelope) 'type value in element xmlcor:ramcod
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcor_aplnumdig", vcor_aplnumdig,"", vXMLRequestEnvelope) 'type value in element xmlcor:aplnumdig
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcor_prporg", vcor_prporg,"", vXMLRequestEnvelope) 'type value in element xmlcor:prporg
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcor_prpnumdig", vcor_prpnumdig,"", vXMLRequestEnvelope) 'type value in element xmlcor:prpnumdig
							Test.TestLog("Consumir o serviço LocalizarPropostaKiprev", "O serviço LocalizarPropostaKiprev deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço LocalizarPropostaKiprev foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService Localizar Propostas
							If p_xmlSoapLeanTest.StatusCode = 200 Then
								Test.TestLog("Verificar o retorno do serviço LocalizarPropostaKiprev", "O serviço LocalizarPropostaKiprev retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço LocalizarPropostaKiprev foi consumido com sucesso", typelog.Passed, , False)
							Else
							Test.TestLog("Verificar o retorno do serviço LocalizarPropostaKiprev", "O serviço LocalizarPropostaKiprev retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço LocalizarPropostaKiprev não foi consumido com sucesso", typelog.Failed, , False)
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
				HandlerError("test_Localizar_Propostas_LocalizarPropostaKiprev.test_Localizar_Propostas_LocalizarPropostaKiprev.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vcor_Cd_filial, vcor_Cd_produto, vcor_Cd_proposta, vcor_succod, vcor_ramcod, vcor_aplnumdig, vcor_prporg, vcor_prpnumdig, vIsOpenSystem As String

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
					vcor_Cd_filial = pc_db.Fieldt("vcor_Cd_filial")
					vcor_Cd_produto = pc_db.Fieldt("vcor_Cd_produto")
					vcor_Cd_proposta = pc_db.Fieldt("vcor_Cd_proposta")
					vcor_succod = pc_db.Fieldt("vcor_succod")
					vcor_ramcod = pc_db.Fieldt("vcor_ramcod")
					vcor_aplnumdig = pc_db.Fieldt("vcor_aplnumdig")
					vcor_prporg = pc_db.Fieldt("vcor_prporg")
					vcor_prpnumdig = pc_db.Fieldt("vcor_prpnumdig")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Localizar_Propostas_LocalizarPropostaKiprev.test_Localizar_Propostas_LocalizarPropostaKiprev.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
