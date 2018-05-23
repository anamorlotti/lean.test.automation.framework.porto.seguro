'**********************************************************************************************************************************************************************
'Create by LeanTest Automation 3.8 in 23/03/2018 13:23:25 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... Protocolo WEB Kiprev
'Functionality:... efetivarProtocoloPropostaPrevidencia
'Master Test:..... No Defined
'TableTest:....... test_Protocolo_WEB_Kiprev_efetivarProtocoloPropostaPrevidencia
'**********************************************************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Protocolo_WEB_Kiprev_efetivarProtocoloPropostaPrevidencia
    Public Class test_Protocolo_WEB_Kiprev_efetivarProtocoloPropostaPrevidencia
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
							vXMLRequestEnvelope = Test.SetValueXMLElement("vorigemSolicitacao", vorigemSolicitacao,"", vXMLRequestEnvelope) 'type value in xml elementorigemSolicitacao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vpropostaSolicitacao", vpropostaSolicitacao,"", vXMLRequestEnvelope) 'type value in xml elementpropostaSolicitacao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoTipoDocumento", vcodigoTipoDocumento,"", vXMLRequestEnvelope) 'type value in xml elementcodigoTipoDocumento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroCertificado", vnumeroCertificado,"", vXMLRequestEnvelope) 'type value in xml elementnumeroCertificado
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoIdentificacaoProtocolo", vcodigoIdentificacaoProtocolo,"", vXMLRequestEnvelope) 'type value in xml elementcodigoIdentificacaoProtocolo
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroOrigemPropostaDestino", vnumeroOrigemPropostaDestino,"", vXMLRequestEnvelope) 'type value in xml elementnumeroOrigemPropostaDestino
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroPropostaDestino", vnumeroPropostaDestino,"", vXMLRequestEnvelope) 'type value in xml elementnumeroPropostaDestino
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoSetorOrigem", vcodigoSetorOrigem,"", vXMLRequestEnvelope) 'type value in xml elementcodigoSetorOrigem
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoFormaPagamento", vcodigoFormaPagamento,"", vXMLRequestEnvelope) 'type value in xml elementcodigoFormaPagamento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vquantidadeParcela", vquantidadeParcela,"", vXMLRequestEnvelope) 'type value in xml elementquantidadeParcela
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoPlano", vcodigoPlano,"", vXMLRequestEnvelope) 'type value in xml elementcodigoPlano
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoRamo", vcodigoRamo,"", vXMLRequestEnvelope) 'type value in xml elementcodigoRamo
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoModalidade", vcodigoModalidade,"", vXMLRequestEnvelope) 'type value in xml elementcodigoModalidade
							vXMLRequestEnvelope = Test.SetValueXMLElement("vvalorContribuicao", vvalorContribuicao,"", vXMLRequestEnvelope) 'type value in xml elementvalorContribuicao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vvalorRisco", vvalorRisco,"", vXMLRequestEnvelope) 'type value in xml elementvalorRisco
							vXMLRequestEnvelope = Test.SetValueXMLElement("vvalorNovoAporte", vvalorNovoAporte,"", vXMLRequestEnvelope) 'type value in xml elementvalorNovoAporte
							vXMLRequestEnvelope = Test.SetValueXMLElement("vorigemProtocolo", vorigemProtocolo,"", vXMLRequestEnvelope) 'type value in xml elementorigemProtocolo
							vXMLRequestEnvelope = Test.SetValueXMLElement("vflagPortabilidade", vflagPortabilidade,"", vXMLRequestEnvelope) 'type value in xml elementflagPortabilidade
							vXMLRequestEnvelope = Test.SetValueXMLElement("vflagTransferencia", vflagTransferencia,"", vXMLRequestEnvelope) 'type value in xml elementflagTransferencia
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoConta", vtipoConta,"", vXMLRequestEnvelope) 'type value in xml elementtipoConta
							vXMLRequestEnvelope = Test.SetValueXMLElement("vdestinoAporte", vdestinoAporte,"", vXMLRequestEnvelope) 'type value in xml elementdestinoAporte
							vXMLRequestEnvelope = Test.SetValueXMLElement("vsusepCorretor", vsusepCorretor,"", vXMLRequestEnvelope) 'type value in xml elementsusepCorretor
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroCpfCnpjSegurado", vnumeroCpfCnpjSegurado,"", vXMLRequestEnvelope) 'type value in xml elementnumeroCpfCnpjSegurado
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnomeSegurado", vnomeSegurado,"", vXMLRequestEnvelope) 'type value in xml elementnomeSegurado
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoPessoaSegurado", vtipoPessoaSegurado,"", vXMLRequestEnvelope) 'type value in xml elementtipoPessoaSegurado
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoSexoSegurado", vtipoSexoSegurado,"", vXMLRequestEnvelope) 'type value in xml elementtipoSexoSegurado
							vXMLRequestEnvelope = Test.SetValueXMLElement("vdataNascimentoSegurado", vdataNascimentoSegurado,"", vXMLRequestEnvelope) 'type value in xml elementdataNascimentoSegurado
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoEstadoCivilSegurado", vtipoEstadoCivilSegurado,"", vXMLRequestEnvelope) 'type value in xml elementtipoEstadoCivilSegurado
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcodigoTipoPagador", vcodigoTipoPagador,"", vXMLRequestEnvelope) 'type value in xml elementcodigoTipoPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroCnpjCpfPagador", vnumeroCnpjCpfPagador,"", vXMLRequestEnvelope) 'type value in xml elementnumeroCnpjCpfPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnomePagador", vnomePagador,"", vXMLRequestEnvelope) 'type value in xml elementnomePagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoPessoaPagador", vtipoPessoaPagador,"", vXMLRequestEnvelope) 'type value in xml elementtipoPessoaPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoSexoPagador", vtipoSexoPagador,"", vXMLRequestEnvelope) 'type value in xml elementtipoSexoPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoEstadoCivilPagador", vtipoEstadoCivilPagador,"", vXMLRequestEnvelope) 'type value in xml elementtipoEstadoCivilPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vdataNascimentoPagador", vdataNascimentoPagador,"", vXMLRequestEnvelope) 'type value in xml elementdataNascimentoPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroCepPagador", vnumeroCepPagador,"", vXMLRequestEnvelope) 'type value in xml elementnumeroCepPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnomeLogradouroPagador", vnomeLogradouroPagador,"", vXMLRequestEnvelope) 'type value in xml elementnomeLogradouroPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroLogradouroPagador", vnumeroLogradouroPagador,"", vXMLRequestEnvelope) 'type value in xml elementnumeroLogradouroPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vcomplementoLogradouroPagador", vcomplementoLogradouroPagador,"", vXMLRequestEnvelope) 'type value in xml elementcomplementoLogradouroPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnomeBairroPagador", vnomeBairroPagador,"", vXMLRequestEnvelope) 'type value in xml elementnomeBairroPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnomeCidadePagador", vnomeCidadePagador,"", vXMLRequestEnvelope) 'type value in xml elementnomeCidadePagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vsiglaUfPagador", vsiglaUfPagador,"", vXMLRequestEnvelope) 'type value in xml elementsiglaUfPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroParcela", vnumeroParcela,"", vXMLRequestEnvelope) 'type value in xml elementnumeroParcela
							vXMLRequestEnvelope = Test.SetValueXMLElement("vtipoRecebimentoDocumento", vtipoRecebimentoDocumento,"", vXMLRequestEnvelope) 'type value in xml elementtipoRecebimentoDocumento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vvalorParcelaRecebimento", vvalorParcelaRecebimento,"", vXMLRequestEnvelope) 'type value in xml elementvalorParcelaRecebimento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroBancoPagador", vnumeroBancoPagador,"", vXMLRequestEnvelope) 'type value in xml elementnumeroBancoPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroAgenciaPagador", vnumeroAgenciaPagador,"", vXMLRequestEnvelope) 'type value in xml elementnumeroAgenciaPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroContaPagador", vnumeroContaPagador,"", vXMLRequestEnvelope) 'type value in xml elementnumeroContaPagador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vnumeroDigitoContaPagador", vnumeroDigitoContaPagador,"", vXMLRequestEnvelope) 'type value in xml elementnumeroDigitoContaPagador
							Test.TestLog("Consumir o serviço efetivarProtocoloPropostaPrevidencia", "O serviço efetivarProtocoloPropostaPrevidencia deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço efetivarProtocoloPropostaPrevidencia foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService Protocolo WEB Kiprev
                            If p_xmlSoapLeanTest.StatusCode = 200 Then
                                Test.TestLog("Verificar o retorno do serviço efetivarProtocoloPropostaPrevidencia", "O serviço efetivarProtocoloPropostaPrevidencia retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço efetivarProtocoloPropostaPrevidencia foi consumido com sucesso", typelog.Passed, , False)
                            Else
                                Test.TestLog("Verificar o retorno do serviço efetivarProtocoloPropostaPrevidencia", "O serviço efetivarProtocoloPropostaPrevidencia retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço efetivarProtocoloPropostaPrevidencia não foi consumido com sucesso", typelog.Failed, , False)
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
		HandlerError("test_Protocolo_WEB_Kiprev_efetivarProtocoloPropostaPrevidencia.test_Protocolo_WEB_Kiprev_efetivarProtocoloPropostaPrevidencia.Run: " & ex.Message)
		Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
		Return False
            End Try
        End Function

        '****************************************************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vorigemSolicitacao, vpropostaSolicitacao, vcodigoTipoDocumento, vnumeroCertificado, vcodigoIdentificacaoProtocolo, vnumeroOrigemPropostaDestino, vnumeroPropostaDestino, vcodigoSetorOrigem, vcodigoFormaPagamento, vquantidadeParcela, vcodigoPlano, vcodigoRamo, vcodigoModalidade, vvalorContribuicao, vvalorRisco, vvalorNovoAporte, vorigemProtocolo, vflagPortabilidade, vflagTransferencia, vtipoConta, vdestinoAporte, vsusepCorretor, vnumeroCpfCnpjSegurado, vnomeSegurado, vtipoPessoaSegurado, vtipoSexoSegurado, vdataNascimentoSegurado, vtipoEstadoCivilSegurado, vcodigoTipoPagador, vnumeroCnpjCpfPagador, vnomePagador, vtipoPessoaPagador, vtipoSexoPagador, vtipoEstadoCivilPagador, vdataNascimentoPagador, vnumeroCepPagador, vnomeLogradouroPagador, vnumeroLogradouroPagador, vcomplementoLogradouroPagador, vnomeBairroPagador, vnomeCidadePagador, vsiglaUfPagador, vnumeroParcela, vtipoRecebimentoDocumento, vvalorParcelaRecebimento, vnumeroBancoPagador, vnumeroAgenciaPagador, vnumeroContaPagador, vnumeroDigitoContaPagador, vIsOpenSystem As String

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
					vorigemSolicitacao = pc_db.Fieldt("vorigemSolicitacao")
					vpropostaSolicitacao = pc_db.Fieldt("vpropostaSolicitacao")
					vcodigoTipoDocumento = pc_db.Fieldt("vcodigoTipoDocumento")
					vnumeroCertificado = pc_db.Fieldt("vnumeroCertificado")
					vcodigoIdentificacaoProtocolo = pc_db.Fieldt("vcodigoIdentificacaoProtocolo")
					vnumeroOrigemPropostaDestino = pc_db.Fieldt("vnumeroOrigemPropostaDestino")
					vnumeroPropostaDestino = pc_db.Fieldt("vnumeroPropostaDestino")
					vcodigoSetorOrigem = pc_db.Fieldt("vcodigoSetorOrigem")
					vcodigoFormaPagamento = pc_db.Fieldt("vcodigoFormaPagamento")
					vquantidadeParcela = pc_db.Fieldt("vquantidadeParcela")
					vcodigoPlano = pc_db.Fieldt("vcodigoPlano")
					vcodigoRamo = pc_db.Fieldt("vcodigoRamo")
					vcodigoModalidade = pc_db.Fieldt("vcodigoModalidade")
					vvalorContribuicao = pc_db.Fieldt("vvalorContribuicao")
					vvalorRisco = pc_db.Fieldt("vvalorRisco")
					vvalorNovoAporte = pc_db.Fieldt("vvalorNovoAporte")
					vorigemProtocolo = pc_db.Fieldt("vorigemProtocolo")
					vflagPortabilidade = pc_db.Fieldt("vflagPortabilidade")
					vflagTransferencia = pc_db.Fieldt("vflagTransferencia")
					vtipoConta = pc_db.Fieldt("vtipoConta")
					vdestinoAporte = pc_db.Fieldt("vdestinoAporte")
					vsusepCorretor = pc_db.Fieldt("vsusepCorretor")
					vnumeroCpfCnpjSegurado = pc_db.Fieldt("vnumeroCpfCnpjSegurado")
					vnomeSegurado = pc_db.Fieldt("vnomeSegurado")
					vtipoPessoaSegurado = pc_db.Fieldt("vtipoPessoaSegurado")
					vtipoSexoSegurado = pc_db.Fieldt("vtipoSexoSegurado")
					vdataNascimentoSegurado = pc_db.Fieldt("vdataNascimentoSegurado")
					vtipoEstadoCivilSegurado = pc_db.Fieldt("vtipoEstadoCivilSegurado")
					vcodigoTipoPagador = pc_db.Fieldt("vcodigoTipoPagador")
					vnumeroCnpjCpfPagador = pc_db.Fieldt("vnumeroCnpjCpfPagador")
					vnomePagador = pc_db.Fieldt("vnomePagador")
					vtipoPessoaPagador = pc_db.Fieldt("vtipoPessoaPagador")
					vtipoSexoPagador = pc_db.Fieldt("vtipoSexoPagador")
					vtipoEstadoCivilPagador = pc_db.Fieldt("vtipoEstadoCivilPagador")
					vdataNascimentoPagador = pc_db.Fieldt("vdataNascimentoPagador")
					vnumeroCepPagador = pc_db.Fieldt("vnumeroCepPagador")
					vnomeLogradouroPagador = pc_db.Fieldt("vnomeLogradouroPagador")
					vnumeroLogradouroPagador = pc_db.Fieldt("vnumeroLogradouroPagador")
					vcomplementoLogradouroPagador = pc_db.Fieldt("vcomplementoLogradouroPagador")
					vnomeBairroPagador = pc_db.Fieldt("vnomeBairroPagador")
					vnomeCidadePagador = pc_db.Fieldt("vnomeCidadePagador")
					vsiglaUfPagador = pc_db.Fieldt("vsiglaUfPagador")
					vnumeroParcela = pc_db.Fieldt("vnumeroParcela")
					vtipoRecebimentoDocumento = pc_db.Fieldt("vtipoRecebimentoDocumento")
					vvalorParcelaRecebimento = pc_db.Fieldt("vvalorParcelaRecebimento")
					vnumeroBancoPagador = pc_db.Fieldt("vnumeroBancoPagador")
					vnumeroAgenciaPagador = pc_db.Fieldt("vnumeroAgenciaPagador")
					vnumeroContaPagador = pc_db.Fieldt("vnumeroContaPagador")
					vnumeroDigitoContaPagador = pc_db.Fieldt("vnumeroDigitoContaPagador")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Protocolo_WEB_Kiprev_efetivarProtocoloPropostaPrevidencia.test_Protocolo_WEB_Kiprev_efetivarProtocoloPropostaPrevidencia.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
