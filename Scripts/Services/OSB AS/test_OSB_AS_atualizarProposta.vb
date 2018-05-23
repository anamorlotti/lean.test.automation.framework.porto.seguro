'**********************************************************************************************************************************************************************
'Create by LeanTest Automation 3.8 in 29/03/2018 15:14:58 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... OSB AS
'Functionality:... atualizarProposta
'Master Test:..... No Defined
'TableTest:....... test_OSB_AS_atualizarProposta
'**********************************************************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_OSB_AS_atualizarProposta
    Public Class test_OSB_AS_atualizarProposta
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "https://osbhmlcorp.portoseguro.brasil:443/CorporativoIntegrationService/FluxoPropostasASIntegrationServiceSoap11V1_0"'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters OSB AS
							vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope OSB AS
							vXMLRequestEnvelope = Test.SetValueXMLElement("vapolice", vapolice,"", vXMLRequestEnvelope) 'type value in element xmlapolice
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoCGC", vcodigoCGC,"", vXMLRequestEnvelope) 'type value in element xmlcodigoCGC
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoConvenio", vcodigoConvenio,"", vXMLRequestEnvelope) 'type value in element xmlcodigoConvenio
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoFormaPagamento", vcodigoFormaPagamento,"", vXMLRequestEnvelope) 'type value in element xmlcodigoFormaPagamento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoModalidade", vcodigoModalidade,"", vXMLRequestEnvelope) 'type value in element xmlcodigoModalidade
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoMotivoDevolucao", vcodigoMotivoDevolucao,"", vXMLRequestEnvelope) 'type value in element xmlcodigoMotivoDevolucao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoRamo", vcodigoRamo,"", vXMLRequestEnvelope) 'type value in element xmlcodigoRamo
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoSucursal", vcodigoSucursal,"", vXMLRequestEnvelope) 'type value in element xmlcodigoSucursal
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoTarefa", vcodigoTarefa,"", vXMLRequestEnvelope) 'type value in element xmlcodigoTarefa
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoUnidade", vcodigoUnidade,"", vXMLRequestEnvelope) 'type value in element xmlcodigoUnidade
							vXMLRequestEnvelope = Test.SetValueXMLElement("vdataEmissaoApolice", vdataEmissaoApolice,"", vXMLRequestEnvelope) 'type value in element xmldataEmissaoApolice
							vXMLRequestEnvelope = Test.SetValueXMLElement("vdigitoCGCCPF", vdigitoCGCCPF,"", vXMLRequestEnvelope) 'type value in element xmldigitoCGCCPF
							vXMLRequestEnvelope = Test.SetValueXMLElement("vempresaFuncionario", vempresaFuncionario,"", vXMLRequestEnvelope) 'type value in element xmlempresaFuncionario
							vXMLRequestEnvelope = Test.SetValueXMLElement("vfimVigencia", vfimVigencia,"", vXMLRequestEnvelope) 'type value in element xmlfimVigencia
							vXMLRequestEnvelope = Test.SetValueXMLElement("vinicioVigencia", vinicioVigencia,"", vXMLRequestEnvelope) 'type value in element xmlinicioVigencia
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmassificados", vmassificados,"", vXMLRequestEnvelope) 'type value in element xmlmassificados
							vXMLRequestEnvelope = Test.SetValueXMLElement("vmatriculaFuncionario", vmatriculaFuncionario,"", vXMLRequestEnvelope) 'type value in element xmlmatriculaFuncionario
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnomeSegurado", vnomeSegurado,"", vXMLRequestEnvelope) 'type value in element xmlnomeSegurado
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroApoliceRenovada", vnumeroApoliceRenovada,"", vXMLRequestEnvelope) 'type value in element xmlnumeroApoliceRenovada
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroCGCCPF", vnumeroCGCCPF,"", vXMLRequestEnvelope) 'type value in element xmlnumeroCGCCPF
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroEndosso", vnumeroEndosso,"", vXMLRequestEnvelope) 'type value in element xmlnumeroEndosso
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroOrigem", vnumeroOrigem,"", vXMLRequestEnvelope) 'type value in element xmlnumeroOrigem
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroProposta", vnumeroProposta,"", vXMLRequestEnvelope) 'type value in element xmlnumeroProposta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroSegurado", vnumeroSegurado,"", vXMLRequestEnvelope) 'type value in element xmlnumeroSegurado
							vXMLRequestEnvelope = Test.SetValueXMLElement("vobservacao", vobservacao,"", vXMLRequestEnvelope) 'type value in element xmlobservacao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vsetor", vsetor,"", vXMLRequestEnvelope) 'type value in element xmlsetor
							vXMLRequestEnvelope = Test.SetValueXMLElement("vsiglaProduto", vsiglaProduto,"", vXMLRequestEnvelope) 'type value in element xmlsiglaProduto
							vXMLRequestEnvelope = Test.SetValueXMLElement("vstatus", vstatus,"", vXMLRequestEnvelope) 'type value in element xmlstatus
							vXMLRequestEnvelope = Test.SetValueXMLElement("vstatusProposta", vstatusProposta,"", vXMLRequestEnvelope) 'type value in element xmlstatusProposta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vsubRamo", vsubRamo,"", vXMLRequestEnvelope) 'type value in element xmlsubRamo
							vXMLRequestEnvelope = Test.SetValueXMLElement("vsusepCorretor", vsusepCorretor,"", vXMLRequestEnvelope) 'type value in element xmlsusepCorretor
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoDocumento", vtipoDocumento,"", vXMLRequestEnvelope) 'type value in element xmltipoDocumento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoEmissao", vtipoEmissao,"", vXMLRequestEnvelope) 'type value in element xmltipoEmissao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoEndosso", vtipoEndosso,"", vXMLRequestEnvelope) 'type value in element xmltipoEndosso
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoFuncionario", vtipoFuncionario,"", vXMLRequestEnvelope) 'type value in element xmltipoFuncionario
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoPessoa", vtipoPessoa,"", vXMLRequestEnvelope) 'type value in element xmltipoPessoa
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoProduto", vtipoProduto,"", vXMLRequestEnvelope) 'type value in element xmltipoProduto
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoProposta", vtipoProposta,"", vXMLRequestEnvelope) 'type value in element xmltipoProposta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vvalorPremio", vvalorPremio,"", vXMLRequestEnvelope) 'type value in element xmlvalorPremio
							Test.TestLog("Consumir o serviço atualizarProposta", "O serviço atualizarProposta deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço atualizarProposta foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService OSB AS
                            If p_xmlSoapLeanTest.StatusCode = 200 Then
                                Test.TestLog("Verificar o retorno do serviço atualizarProposta", "O serviço atualizarProposta retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço atualizarProposta foi consumido com sucesso", typelog.Passed, , False)
                            Else
                                Test.TestLog("Verificar o retorno do serviço atualizarProposta", "O serviço atualizarProposta retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço atualizarProposta não foi consumido com sucesso", typelog.Failed, , False)
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
		HandlerError("test_OSB_AS_atualizarProposta.test_OSB_AS_atualizarProposta.Run: " & ex.Message)
		Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
		Return False
            End Try
        End Function

        '****************************************************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vapolice, vcodigoCGC, vcodigoConvenio, vcodigoFormaPagamento, vcodigoModalidade, vcodigoMotivoDevolucao, vcodigoRamo, vcodigoSucursal, vcodigoTarefa, vcodigoUnidade, vdataEmissaoApolice, vdigitoCGCCPF, vempresaFuncionario, vfimVigencia, vinicioVigencia, vmassificados, vmatriculaFuncionario, vnomeSegurado, vnumeroApoliceRenovada, vnumeroCGCCPF, vnumeroEndosso, vnumeroOrigem, vnumeroProposta, vnumeroSegurado, vobservacao, vsetor, vsiglaProduto, vstatus, vstatusProposta, vsubRamo, vsusepCorretor, vtipoDocumento, vtipoEmissao, vtipoEndosso, vtipoFuncionario, vtipoPessoa, vtipoProduto, vtipoProposta, vvalorPremio, vIsOpenSystem As String

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
					vapolice = pc_db.Fieldt("vapolice")
					vcodigoCGC = pc_db.Fieldt("vcodigoCGC")
					vcodigoConvenio = pc_db.Fieldt("vcodigoConvenio")
					vcodigoFormaPagamento = pc_db.Fieldt("vcodigoFormaPagamento")
					vcodigoModalidade = pc_db.Fieldt("vcodigoModalidade")
					vcodigoMotivoDevolucao = pc_db.Fieldt("vcodigoMotivoDevolucao")
					vcodigoRamo = pc_db.Fieldt("vcodigoRamo")
					vcodigoSucursal = pc_db.Fieldt("vcodigoSucursal")
					vcodigoTarefa = pc_db.Fieldt("vcodigoTarefa")
					vcodigoUnidade = pc_db.Fieldt("vcodigoUnidade")
					vdataEmissaoApolice = pc_db.Fieldt("vdataEmissaoApolice")
					vdigitoCGCCPF = pc_db.Fieldt("vdigitoCGCCPF")
					vempresaFuncionario = pc_db.Fieldt("vempresaFuncionario")
					vfimVigencia = pc_db.Fieldt("vfimVigencia")
					vinicioVigencia = pc_db.Fieldt("vinicioVigencia")
					vmassificados = pc_db.Fieldt("vmassificados")
					vmatriculaFuncionario = pc_db.Fieldt("vmatriculaFuncionario")
					vnomeSegurado = pc_db.Fieldt("vnomeSegurado")
					vnumeroApoliceRenovada = pc_db.Fieldt("vnumeroApoliceRenovada")
					vnumeroCGCCPF = pc_db.Fieldt("vnumeroCGCCPF")
					vnumeroEndosso = pc_db.Fieldt("vnumeroEndosso")
					vnumeroOrigem = pc_db.Fieldt("vnumeroOrigem")
					vnumeroProposta = pc_db.Fieldt("vnumeroProposta")
					vnumeroSegurado = pc_db.Fieldt("vnumeroSegurado")
					vobservacao = pc_db.Fieldt("vobservacao")
					vsetor = pc_db.Fieldt("vsetor")
					vsiglaProduto = pc_db.Fieldt("vsiglaProduto")
					vstatus = pc_db.Fieldt("vstatus")
					vstatusProposta = pc_db.Fieldt("vstatusProposta")
					vsubRamo = pc_db.Fieldt("vsubRamo")
					vsusepCorretor = pc_db.Fieldt("vsusepCorretor")
					vtipoDocumento = pc_db.Fieldt("vtipoDocumento")
					vtipoEmissao = pc_db.Fieldt("vtipoEmissao")
					vtipoEndosso = pc_db.Fieldt("vtipoEndosso")
					vtipoFuncionario = pc_db.Fieldt("vtipoFuncionario")
					vtipoPessoa = pc_db.Fieldt("vtipoPessoa")
					vtipoProduto = pc_db.Fieldt("vtipoProduto")
					vtipoProposta = pc_db.Fieldt("vtipoProposta")
					vvalorPremio = pc_db.Fieldt("vvalorPremio")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_OSB_AS_atualizarProposta.test_OSB_AS_atualizarProposta.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
