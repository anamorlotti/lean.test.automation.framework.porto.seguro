'**********************************************************************************************************************************************************************
'Create by LeanTest Automation 3.8 in 22/03/2018 15:40:38 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... Protocolo WEB Kiprev
'Functionality:... devolverProtocoloPrevidencia
'Master Test:..... No Defined
'TableTest:....... test_Protocolo_WEB_Kiprev_devolverProtocoloPrevidencia
'**********************************************************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Protocolo_WEB_Kiprev_devolverProtocoloPrevidencia
    Public Class test_Protocolo_WEB_Kiprev_devolverProtocoloPrevidencia
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "http://kiprevwebhml:80/ki-ws-protocolo/ProtocoloWS"'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters Protocolo WEB Kiprev
							vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope Protocolo WEB Kiprev
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoDocumento", vtipoDocumento,"", vXMLRequestEnvelope) 'type value in xml elementtipoDocumento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vorigemSolicitacao", vorigemSolicitacao,"", vXMLRequestEnvelope) 'type value in xml elementorigemSolicitacao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vpropostaSolicitacao", vpropostaSolicitacao,"", vXMLRequestEnvelope) 'type value in xml elementpropostaSolicitacao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vsusepCorretor", vsusepCorretor,"", vXMLRequestEnvelope) 'type value in xml elementsusepCorretor
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmatriculaDevolucao", vmatriculaDevolucao,"", vXMLRequestEnvelope) 'type value in xml elementmatriculaDevolucao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoDevolucao", vcodigoDevolucao,"", vXMLRequestEnvelope) 'type value in xml elementcodigoDevolucao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmotivoDevolucao", vmotivoDevolucao,"", vXMLRequestEnvelope) 'type value in xml elementmotivoDevolucao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoSetorOrigem", vcodigoSetorOrigem,"", vXMLRequestEnvelope) 'type value in xml elementcodigoSetorOrigem
							Test.TestLog("Consumir o serviço devolverProtocoloPrevidencia", "O serviço devolverProtocoloPrevidencia deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço devolverProtocoloPrevidencia foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService Protocolo WEB Kiprev
                            If p_xmlSoapLeanTest.StatusCode = 200 Then
                                Test.TestLog("Verificar o retorno do serviço devolverProtocoloPrevidencia", "O serviço devolverProtocoloPrevidencia retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço devolverProtocoloPrevidencia foi consumido com sucesso", typelog.Passed, , False)
                            Else
                                Test.TestLog("Verificar o retorno do serviço devolverProtocoloPrevidencia", "O serviço devolverProtocoloPrevidencia retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço devolverProtocoloPrevidencia não foi consumido com sucesso", typelog.Failed, , False)
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
		HandlerError("test_Protocolo_WEB_Kiprev_devolverProtocoloPrevidencia.test_Protocolo_WEB_Kiprev_devolverProtocoloPrevidencia.Run: " & ex.Message)
		Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
		Return False
            End Try
        End Function

        '****************************************************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vtipoDocumento, vorigemSolicitacao, vpropostaSolicitacao, vsusepCorretor, vmatriculaDevolucao, vcodigoDevolucao, vmotivoDevolucao, vcodigoSetorOrigem, vIsOpenSystem As String

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
					vtipoDocumento = pc_db.Fieldt("vtipoDocumento")
					vorigemSolicitacao = pc_db.Fieldt("vorigemSolicitacao")
					vpropostaSolicitacao = pc_db.Fieldt("vpropostaSolicitacao")
					vsusepCorretor = pc_db.Fieldt("vsusepCorretor")
					vmatriculaDevolucao = pc_db.Fieldt("vmatriculaDevolucao")
					vcodigoDevolucao = pc_db.Fieldt("vcodigoDevolucao")
					vmotivoDevolucao = pc_db.Fieldt("vmotivoDevolucao")
					vcodigoSetorOrigem = pc_db.Fieldt("vcodigoSetorOrigem")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Protocolo_WEB_Kiprev_devolverProtocoloPrevidencia.test_Protocolo_WEB_Kiprev_devolverProtocoloPrevidencia.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
