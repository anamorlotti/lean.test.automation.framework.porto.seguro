'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 11/04/2018 10:59:15 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to WebService
'Environment:..... Automation Project
'Application...... OSB BCP
'Functionality:... adicionarPessoaProduto
'Master Test:..... No Defined
'TableTest:....... test_OSB_BCP_adicionarPessoaProduto
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_OSB_BCP_adicionarPessoaProduto
    Public Class test_OSB_BCP_adicionarPessoaProduto
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "  http://osbhmlcrm.portoseguro.brasil:80/PessoaProdutoIntegrationService/Proxy_Services/PessoaProdutoIntegrationServiceSoap11V1_1 "'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vXMLParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLParameters") 'get XMLParameters OSB BCP
							vXMLRequestEnvelope = Test.GetValueOutput(p_IDTest, p_TableTest, "vXMLRequestEnvelope") 'get XMLRequestEnvelope OSB BCP
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_codigoOrigemMovimento", vv11_codigoOrigemMovimento,"", vXMLRequestEnvelope) 'type value in element xmlv11:codigoOrigemMovimento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_grupoXML", vv11_grupoXML,"", vXMLRequestEnvelope) 'type value in element xmlv11:grupoXML
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_dataHoraMoviento", vv11_dataHoraMoviento,"", vXMLRequestEnvelope) 'type value in element xmlv11:dataHoraMoviento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_nome", vv11_nome,"", vXMLRequestEnvelope) 'type value in element xmlv11:nome
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_numeroCnpjCpf", vv12_numeroCnpjCpf,"", vXMLRequestEnvelope) 'type value in element xmlv12:numeroCnpjCpf
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_digitoCnpjCpf", vv12_digitoCnpjCpf,"", vXMLRequestEnvelope) 'type value in element xmlv12:digitoCnpjCpf
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_titularDoc", vv11_titularDoc,"", vXMLRequestEnvelope) 'type value in element xmlv11:titularDoc
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_papel", vv11_papel,"", vXMLRequestEnvelope) 'type value in element xmlv11:papel
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_origem", vv11_origem,"", vXMLRequestEnvelope) 'type value in element xmlv11:origem
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_tipoPessoa", vv11_tipoPessoa,"", vXMLRequestEnvelope) 'type value in element xmlv11:tipoPessoa
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_tipoPrestador", vv11_tipoPrestador,"", vXMLRequestEnvelope) 'type value in element xmlv11:tipoPrestador
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_sexo", vv11_sexo,"", vXMLRequestEnvelope) 'type value in element xmlv11:sexo
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_nascimento", vv11_nascimento,"", vXMLRequestEnvelope) 'type value in element xmlv11:nascimento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_estCivil", vv11_estCivil,"", vXMLRequestEnvelope) 'type value in element xmlv11:estCivil
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_situacao", vv11_situacao,"", vXMLRequestEnvelope) 'type value in element xmlv11:situacao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_codEmpresa", vv11_codEmpresa,"", vXMLRequestEnvelope) 'type value in element xmlv11:codEmpresa
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_fxRendaCod", vv11_fxRendaCod,"", vXMLRequestEnvelope) 'type value in element xmlv11:fxRendaCod
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_emailTip", vv11_emailTip,"", vXMLRequestEnvelope) 'type value in element xmlv11:emailTip
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_email", vv11_email,"", vXMLRequestEnvelope) 'type value in element xmlv11:email
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_emailFlgOpt", vv11_emailFlgOpt,"", vXMLRequestEnvelope) 'type value in element xmlv11:emailFlgOpt
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_tipoLogradouro", vv12_tipoLogradouro,"", vXMLRequestEnvelope) 'type value in element xmlv12:tipoLogradouro
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_logradouro", vv12_logradouro,"", vXMLRequestEnvelope) 'type value in element xmlv12:logradouro
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_numeroLogradouro", vv12_numeroLogradouro,"", vXMLRequestEnvelope) 'type value in element xmlv12:numeroLogradouro
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_bairro", vv12_bairro,"", vXMLRequestEnvelope) 'type value in element xmlv12:bairro
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_cidade", vv12_cidade,"", vXMLRequestEnvelope) 'type value in element xmlv12:cidade
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_uf", vv12_uf,"", vXMLRequestEnvelope) 'type value in element xmlv12:uf
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_cep", vv12_cep,"", vXMLRequestEnvelope) 'type value in element xmlv12:cep
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_complementoCep", vv12_complementoCep,"", vXMLRequestEnvelope) 'type value in element xmlv12:complementoCep
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_siglPais", vv12_siglPais,"", vXMLRequestEnvelope) 'type value in element xmlv12:siglPais
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_flagAutorizacaoPropaganda", vv12_flagAutorizacaoPropaganda,"", vXMLRequestEnvelope) 'type value in element xmlv12:flagAutorizacaoPropaganda
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_codigoOrigemMovimento", vv12_codigoOrigemMovimento,"", vXMLRequestEnvelope) 'type value in element xmlv12:codigoOrigemMovimento
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_siglPais", vv11_siglPais,"", vXMLRequestEnvelope) 'type value in element xmlv11:siglPais
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_enderFlgOpt", vv11_enderFlgOpt,"", vXMLRequestEnvelope) 'type value in element xmlv11:enderFlgOpt
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_foneTip", vv11_foneTip,"", vXMLRequestEnvelope) 'type value in element xmlv11:foneTip
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_foneDDD", vv11_foneDDD,"", vXMLRequestEnvelope) 'type value in element xmlv11:foneDDD
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_foneNum", vv11_foneNum,"", vXMLRequestEnvelope) 'type value in element xmlv11:foneNum
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_flagPep", vv11_flagPep,"", vXMLRequestEnvelope) 'type value in element xmlv11:flagPep
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_flagOpcao", vv12_flagOpcao,"", vXMLRequestEnvelope) 'type value in element xmlv12:flagOpcao
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_numero", vv11_numero,"", vXMLRequestEnvelope) 'type value in element xmlv11:numero
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_alterTip", vv11_alterTip,"", vXMLRequestEnvelope) 'type value in element xmlv11:alterTip
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_edsTip", vv11_edsTip,"", vXMLRequestEnvelope) 'type value in element xmlv11:edsTip
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_dtAtuStatus", vv11_dtAtuStatus,"", vXMLRequestEnvelope) 'type value in element xmlv11:dtAtuStatus
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_dtEmiss", vv11_dtEmiss,"", vXMLRequestEnvelope) 'type value in element xmlv11:dtEmiss
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_formaPagto", vv11_formaPagto,"", vXMLRequestEnvelope) 'type value in element xmlv11:formaPagto
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_origemPrp", vv11_origemPrp,"", vXMLRequestEnvelope) 'type value in element xmlv11:origemPrp
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_numPrp", vv11_numPrp,"", vXMLRequestEnvelope) 'type value in element xmlv11:numPrp
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_tipoLogradouro_1", vv12_tipoLogradouro_1,"", vXMLRequestEnvelope) 'type value in element xmlv12:tipoLogradouro_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_logradouro_1", vv12_logradouro_1,"", vXMLRequestEnvelope) 'type value in element xmlv12:logradouro_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_numeroLogradouro_1", vv12_numeroLogradouro_1,"", vXMLRequestEnvelope) 'type value in element xmlv12:numeroLogradouro_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_bairro_1", vv12_bairro_1,"", vXMLRequestEnvelope) 'type value in element xmlv12:bairro_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_cidade_1", vv12_cidade_1,"", vXMLRequestEnvelope) 'type value in element xmlv12:cidade_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_uf_1", vv12_uf_1,"", vXMLRequestEnvelope) 'type value in element xmlv12:uf_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_cep_1", vv12_cep_1,"", vXMLRequestEnvelope) 'type value in element xmlv12:cep_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_complementoCep_1", vv12_complementoCep_1,"", vXMLRequestEnvelope) 'type value in element xmlv12:complementoCep_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_siglaPais", vv12_siglaPais,"", vXMLRequestEnvelope) 'type value in element xmlv12:siglaPais
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_flagAutorizacaoPropaganda_1", vv12_flagAutorizacaoPropaganda_1,"", vXMLRequestEnvelope) 'type value in element xmlv12:flagAutorizacaoPropaganda_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_codigoOrigemMovimento_1", vv12_codigoOrigemMovimento_1,"", vXMLRequestEnvelope) 'type value in element xmlv12:codigoOrigemMovimento_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_siglPais_1", vv11_siglPais_1,"", vXMLRequestEnvelope) 'type value in element xmlv11:siglPais_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_enderFlgOpt_1", vv11_enderFlgOpt_1,"", vXMLRequestEnvelope) 'type value in element xmlv11:enderFlgOpt_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_empCod", vv11_empCod,"", vXMLRequestEnvelope) 'type value in element xmlv11:empCod
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_dtAtuStatus_1", vv11_dtAtuStatus_1,"", vXMLRequestEnvelope) 'type value in element xmlv11:dtAtuStatus_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_papel_1", vv11_papel_1,"", vXMLRequestEnvelope) 'type value in element xmlv11:papel_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_tipoLogradouro_2", vv12_tipoLogradouro_2,"", vXMLRequestEnvelope) 'type value in element xmlv12:tipoLogradouro_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_logradouro_2", vv12_logradouro_2,"", vXMLRequestEnvelope) 'type value in element xmlv12:logradouro_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_numeroLogradouro_2", vv12_numeroLogradouro_2,"", vXMLRequestEnvelope) 'type value in element xmlv12:numeroLogradouro_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_bairro_2", vv12_bairro_2,"", vXMLRequestEnvelope) 'type value in element xmlv12:bairro_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_cidade_2", vv12_cidade_2,"", vXMLRequestEnvelope) 'type value in element xmlv12:cidade_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_uf_2", vv12_uf_2,"", vXMLRequestEnvelope) 'type value in element xmlv12:uf_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_cep_2", vv12_cep_2,"", vXMLRequestEnvelope) 'type value in element xmlv12:cep_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_complementoCep_2", vv12_complementoCep_2,"", vXMLRequestEnvelope) 'type value in element xmlv12:complementoCep_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_siglaPais_1", vv12_siglaPais_1,"", vXMLRequestEnvelope) 'type value in element xmlv12:siglaPais_1
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_flagAutorizacaoPropaganda_2", vv12_flagAutorizacaoPropaganda_2,"", vXMLRequestEnvelope) 'type value in element xmlv12:flagAutorizacaoPropaganda_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv12_codigoOrigemMovimento_2", vv12_codigoOrigemMovimento_2,"", vXMLRequestEnvelope) 'type value in element xmlv12:codigoOrigemMovimento_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_siglPais_2", vv11_siglPais_2,"", vXMLRequestEnvelope) 'type value in element xmlv11:siglPais_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_enderFlgOpt_2", vv11_enderFlgOpt_2,"", vXMLRequestEnvelope) 'type value in element xmlv11:enderFlgOpt_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_foneTip_2", vv11_foneTip_2,"", vXMLRequestEnvelope) 'type value in element xmlv11:foneTip_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_foneDDD_2", vv11_foneDDD_2,"", vXMLRequestEnvelope) 'type value in element xmlv11:foneDDD_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_foneNum_2", vv11_foneNum_2,"", vXMLRequestEnvelope) 'type value in element xmlv11:foneNum_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_susep", vv11_susep,"", vXMLRequestEnvelope) 'type value in element xmlv11:susep
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_dtEmiss_2", vv11_dtEmiss_2,"", vXMLRequestEnvelope) 'type value in element xmlv11:dtEmiss_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_emailTip_2", vv11_emailTip_2,"", vXMLRequestEnvelope) 'type value in element xmlv11:emailTip_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_email_2", vv11_email_2,"", vXMLRequestEnvelope) 'type value in element xmlv11:email_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_emailFlgOpt_2", vv11_emailFlgOpt_2,"", vXMLRequestEnvelope) 'type value in element xmlv11:emailFlgOpt_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_origemPrp_2", vv11_origemPrp_2,"", vXMLRequestEnvelope) 'type value in element xmlv11:origemPrp_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_numPrp_2", vv11_numPrp_2,"", vXMLRequestEnvelope) 'type value in element xmlv11:numPrp_2
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_segCod", vv11_segCod,"", vXMLRequestEnvelope) 'type value in element xmlv11:segCod
							vXMLRequestEnvelope = Test.SetValueXMLElement("vv11_canalVend", vv11_canalVend,"", vXMLRequestEnvelope) 'type value in element xmlv11:canalVend
							Test.TestLog("Consumir o serviço adicionarPessoaProduto", "O serviço adicionarPessoaProduto deve ser consumido com sucesso usando os valores " & vbCrLf & vXMLRequestEnvelope, "O serviço adicionarPessoaProduto foi consumido com sucesso", typelog.Passed, , False)
							p_xmlSoapLeanTest = Test.CallWebService(p_pathUrlApp, vXMLRequestEnvelope, vXMLParameters, p_TypeExecution) 'CallWebService OSB BCP
							If p_xmlSoapLeanTest.StatusCode = 200 Then
								Test.TestLog("Verificar o retorno do serviço adicionarPessoaProduto", "O serviço adicionarPessoaProduto retornou com sucesso usando os valores " & vbCrLf & p_xmlSoapLeanTest.InnerXML, "O serviço adicionarPessoaProduto foi consumido com sucesso", typelog.Passed, , False)
							Else
							Test.TestLog("Verificar o retorno do serviço adicionarPessoaProduto", "O serviço adicionarPessoaProduto retornou a mensagem de erro: " & vbCrLf & p_xmlSoapLeanTest.ResponseError, "O serviço adicionarPessoaProduto não foi consumido com sucesso", typelog.Failed, , False)
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
				HandlerError("test_OSB_BCP_adicionarPessoaProduto.test_OSB_BCP_adicionarPessoaProduto.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vXMLParameters,vXMLRequestEnvelope,vv11_codigoOrigemMovimento, vv11_grupoXML, vv11_dataHoraMoviento, vv11_nome, vv12_numeroCnpjCpf, vv12_digitoCnpjCpf, vv11_titularDoc, vv11_papel, vv11_origem, vv11_tipoPessoa, vv11_tipoPrestador, vv11_sexo, vv11_nascimento, vv11_estCivil, vv11_situacao, vv11_codEmpresa, vv11_fxRendaCod, vv11_emailTip, vv11_email, vv11_emailFlgOpt, vv12_tipoLogradouro, vv12_logradouro, vv12_numeroLogradouro, vv12_bairro, vv12_cidade, vv12_uf, vv12_cep, vv12_complementoCep, vv12_siglPais, vv12_flagAutorizacaoPropaganda, vv12_codigoOrigemMovimento, vv11_siglPais, vv11_enderFlgOpt, vv11_foneTip, vv11_foneDDD, vv11_foneNum, vv11_flagPep, vv12_flagOpcao, vv11_numero, vv11_alterTip, vv11_edsTip, vv11_dtAtuStatus, vv11_dtEmiss, vv11_formaPagto, vv11_origemPrp, vv11_numPrp, vv12_tipoLogradouro_1, vv12_logradouro_1, vv12_numeroLogradouro_1, vv12_bairro_1, vv12_cidade_1, vv12_uf_1, vv12_cep_1, vv12_complementoCep_1, vv12_siglaPais, vv12_flagAutorizacaoPropaganda_1, vv12_codigoOrigemMovimento_1, vv11_siglPais_1, vv11_enderFlgOpt_1, vv11_empCod, vv11_dtAtuStatus_1, vv11_papel_1, vv12_tipoLogradouro_2, vv12_logradouro_2, vv12_numeroLogradouro_2, vv12_bairro_2, vv12_cidade_2, vv12_uf_2, vv12_cep_2, vv12_complementoCep_2, vv12_siglaPais_1, vv12_flagAutorizacaoPropaganda_2, vv12_codigoOrigemMovimento_2, vv11_siglPais_2, vv11_enderFlgOpt_2, vv11_foneTip_2, vv11_foneDDD_2, vv11_foneNum_2, vv11_susep, vv11_dtEmiss_2, vv11_emailTip_2, vv11_email_2, vv11_emailFlgOpt_2, vv11_origemPrp_2, vv11_numPrp_2, vv11_segCod, vv11_canalVend, vIsOpenSystem As String

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
					vv11_codigoOrigemMovimento = pc_db.Fieldt("vv11_codigoOrigemMovimento")
					vv11_grupoXML = pc_db.Fieldt("vv11_grupoXML")
					vv11_dataHoraMoviento = pc_db.Fieldt("vv11_dataHoraMoviento")
					vv11_nome = pc_db.Fieldt("vv11_nome")
					vv12_numeroCnpjCpf = pc_db.Fieldt("vv12_numeroCnpjCpf")
					vv12_digitoCnpjCpf = pc_db.Fieldt("vv12_digitoCnpjCpf")
					vv11_titularDoc = pc_db.Fieldt("vv11_titularDoc")
					vv11_papel = pc_db.Fieldt("vv11_papel")
					vv11_origem = pc_db.Fieldt("vv11_origem")
					vv11_tipoPessoa = pc_db.Fieldt("vv11_tipoPessoa")
					vv11_tipoPrestador = pc_db.Fieldt("vv11_tipoPrestador")
					vv11_sexo = pc_db.Fieldt("vv11_sexo")
					vv11_nascimento = pc_db.Fieldt("vv11_nascimento")
					vv11_estCivil = pc_db.Fieldt("vv11_estCivil")
					vv11_situacao = pc_db.Fieldt("vv11_situacao")
					vv11_codEmpresa = pc_db.Fieldt("vv11_codEmpresa")
					vv11_fxRendaCod = pc_db.Fieldt("vv11_fxRendaCod")
					vv11_emailTip = pc_db.Fieldt("vv11_emailTip")
					vv11_email = pc_db.Fieldt("vv11_email")
					vv11_emailFlgOpt = pc_db.Fieldt("vv11_emailFlgOpt")
					vv12_tipoLogradouro = pc_db.Fieldt("vv12_tipoLogradouro")
					vv12_logradouro = pc_db.Fieldt("vv12_logradouro")
					vv12_numeroLogradouro = pc_db.Fieldt("vv12_numeroLogradouro")
					vv12_bairro = pc_db.Fieldt("vv12_bairro")
					vv12_cidade = pc_db.Fieldt("vv12_cidade")
					vv12_uf = pc_db.Fieldt("vv12_uf")
					vv12_cep = pc_db.Fieldt("vv12_cep")
					vv12_complementoCep = pc_db.Fieldt("vv12_complementoCep")
					vv12_siglPais = pc_db.Fieldt("vv12_siglPais")
					vv12_flagAutorizacaoPropaganda = pc_db.Fieldt("vv12_flagAutorizacaoPropaganda")
					vv12_codigoOrigemMovimento = pc_db.Fieldt("vv12_codigoOrigemMovimento")
					vv11_siglPais = pc_db.Fieldt("vv11_siglPais")
					vv11_enderFlgOpt = pc_db.Fieldt("vv11_enderFlgOpt")
					vv11_foneTip = pc_db.Fieldt("vv11_foneTip")
					vv11_foneDDD = pc_db.Fieldt("vv11_foneDDD")
					vv11_foneNum = pc_db.Fieldt("vv11_foneNum")
					vv11_flagPep = pc_db.Fieldt("vv11_flagPep")
					vv12_flagOpcao = pc_db.Fieldt("vv12_flagOpcao")
					vv11_numero = pc_db.Fieldt("vv11_numero")
					vv11_alterTip = pc_db.Fieldt("vv11_alterTip")
					vv11_edsTip = pc_db.Fieldt("vv11_edsTip")
					vv11_dtAtuStatus = pc_db.Fieldt("vv11_dtAtuStatus")
					vv11_dtEmiss = pc_db.Fieldt("vv11_dtEmiss")
					vv11_formaPagto = pc_db.Fieldt("vv11_formaPagto")
					vv11_origemPrp = pc_db.Fieldt("vv11_origemPrp")
					vv11_numPrp = pc_db.Fieldt("vv11_numPrp")
					vv12_tipoLogradouro_1 = pc_db.Fieldt("vv12_tipoLogradouro_1")
					vv12_logradouro_1 = pc_db.Fieldt("vv12_logradouro_1")
					vv12_numeroLogradouro_1 = pc_db.Fieldt("vv12_numeroLogradouro_1")
					vv12_bairro_1 = pc_db.Fieldt("vv12_bairro_1")
					vv12_cidade_1 = pc_db.Fieldt("vv12_cidade_1")
					vv12_uf_1 = pc_db.Fieldt("vv12_uf_1")
					vv12_cep_1 = pc_db.Fieldt("vv12_cep_1")
					vv12_complementoCep_1 = pc_db.Fieldt("vv12_complementoCep_1")
					vv12_siglaPais = pc_db.Fieldt("vv12_siglaPais")
					vv12_flagAutorizacaoPropaganda_1 = pc_db.Fieldt("vv12_flagAutorizacaoPropaganda_1")
					vv12_codigoOrigemMovimento_1 = pc_db.Fieldt("vv12_codigoOrigemMovimento_1")
					vv11_siglPais_1 = pc_db.Fieldt("vv11_siglPais_1")
					vv11_enderFlgOpt_1 = pc_db.Fieldt("vv11_enderFlgOpt_1")
					vv11_empCod = pc_db.Fieldt("vv11_empCod")
					vv11_dtAtuStatus_1 = pc_db.Fieldt("vv11_dtAtuStatus_1")
					vv11_papel_1 = pc_db.Fieldt("vv11_papel_1")
					vv12_tipoLogradouro_2 = pc_db.Fieldt("vv12_tipoLogradouro_2")
					vv12_logradouro_2 = pc_db.Fieldt("vv12_logradouro_2")
					vv12_numeroLogradouro_2 = pc_db.Fieldt("vv12_numeroLogradouro_2")
					vv12_bairro_2 = pc_db.Fieldt("vv12_bairro_2")
					vv12_cidade_2 = pc_db.Fieldt("vv12_cidade_2")
					vv12_uf_2 = pc_db.Fieldt("vv12_uf_2")
					vv12_cep_2 = pc_db.Fieldt("vv12_cep_2")
					vv12_complementoCep_2 = pc_db.Fieldt("vv12_complementoCep_2")
					vv12_siglaPais_1 = pc_db.Fieldt("vv12_siglaPais_1")
					vv12_flagAutorizacaoPropaganda_2 = pc_db.Fieldt("vv12_flagAutorizacaoPropaganda_2")
					vv12_codigoOrigemMovimento_2 = pc_db.Fieldt("vv12_codigoOrigemMovimento_2")
					vv11_siglPais_2 = pc_db.Fieldt("vv11_siglPais_2")
					vv11_enderFlgOpt_2 = pc_db.Fieldt("vv11_enderFlgOpt_2")
					vv11_foneTip_2 = pc_db.Fieldt("vv11_foneTip_2")
					vv11_foneDDD_2 = pc_db.Fieldt("vv11_foneDDD_2")
					vv11_foneNum_2 = pc_db.Fieldt("vv11_foneNum_2")
					vv11_susep = pc_db.Fieldt("vv11_susep")
					vv11_dtEmiss_2 = pc_db.Fieldt("vv11_dtEmiss_2")
					vv11_emailTip_2 = pc_db.Fieldt("vv11_emailTip_2")
					vv11_email_2 = pc_db.Fieldt("vv11_email_2")
					vv11_emailFlgOpt_2 = pc_db.Fieldt("vv11_emailFlgOpt_2")
					vv11_origemPrp_2 = pc_db.Fieldt("vv11_origemPrp_2")
					vv11_numPrp_2 = pc_db.Fieldt("vv11_numPrp_2")
					vv11_segCod = pc_db.Fieldt("vv11_segCod")
					vv11_canalVend = pc_db.Fieldt("vv11_canalVend")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_OSB_BCP_adicionarPessoaProduto.test_OSB_BCP_adicionarPessoaProduto.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
