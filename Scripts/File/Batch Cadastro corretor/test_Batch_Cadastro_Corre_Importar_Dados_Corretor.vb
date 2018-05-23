'**********************************************************************************************************************************************************************
'Create by LeanTest Automation 3.8 in 21/05/2018 14:22:35 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to File
'Environment:..... Automation Project
'Application...... Batch Cadastro Corre
'Functionality:... Importar Dados Corretor
'Master Test:..... No Defined
'TableTest:....... test_Batch_Cadastro_Corre_Importar_Dados_Corretor
'**********************************************************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Batch_Cadastro_Corre_Importar_Dados_Corretor
    Public Class test_Batch_Cadastro_Corre_Importar_Dados_Corretor
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\Scripts\File\Batch Cadastro  Corretor\corddmm.txt"'xml.Read(""urlLocal"") 'Create urlLocal element in XML
								Test.TestLog("Abrir o arquivo Batch Cadastro Corre", "O arquivo  Importar Dados Corretor deve ser aberto para leitura", "O arquivo Importar Dados Corretor foi aberto com sucesso", typelog.Passed, , False)
							End If
							Test.ValidadeFileField(p_pathUrlApp, "0;;1;1;Numeric", "Codigo_Arquivo",vCodigo_Arquivo) 'validate filedCódigo Arquivo
							Test.ValidadeFileField(p_pathUrlApp, "0;;3;8;Others", "Cd_susep",vCd_susep) 'validate filedCd_susep
							Test.ValidadeFileField(p_pathUrlApp, "0;;9;48;Others", "Nm_corretor",vNm_corretor) 'validate filedNm_corretor
							Test.ValidadeFileField(p_pathUrlApp, "0;;49;60;Numeric", "Nr_cgc_cpf",vNr_cgc_cpf) 'validate filedNr_cgc_cpf
							Test.ValidadeFileField(p_pathUrlApp, "0;;61;64;Numeric", "Nr_ordem_cgc",vNr_ordem_cgc) 'validate filedNr_ordem_cgc
							Test.ValidadeFileField(p_pathUrlApp, "0;;65;66;Numeric", "Nr_digito_cgc_cpf",vNr_digito_cgc_cpf) 'validate filedNr_digito_cgc_cpf
							Test.ValidadeFileField(p_pathUrlApp, "0;;68;107;Others", "Nm_endereco",vNm_endereco) 'validate filedNm_endereco
							Test.ValidadeFileField(p_pathUrlApp, "0;;108;113;Others", "Nr_endereco",vNr_endereco) 'validate filedNr_endereco
							Test.ValidadeFileField(p_pathUrlApp, "0;;114;128;Others", "Nm_complemento",vNm_complemento) 'validate filedNm_complemento
							Test.ValidadeFileField(p_pathUrlApp, "0;;129;148;Others", "Nm_bairro",vNm_bairro) 'validate filedNm_bairro
							Test.ValidadeFileField(p_pathUrlApp, "0;;149;168;Others", "Nm_cidade",vNm_cidade) 'validate filedNm_cidade
							Test.ValidadeFileField(p_pathUrlApp, "0;;169;170;Alphabetic", "Cd_uf",vCd_uf) 'validate filedCd_uf
							Test.ValidadeFileField(p_pathUrlApp, "0;;171;178;Others", "Cd_cep",vCd_cep) 'validate filedCd_cep
							Test.ValidadeFileField(p_pathUrlApp, "0;;180;182;Alphanumeric", "Nr_tp_enderecamento",vNr_tp_enderecamento) 'validate filedNr_tp_enderecamento
							Test.ValidadeFileField(p_pathUrlApp, "0;;183;187;AOthers", "Cd_escritorio_negoci",vCd_escritorio_negoci) 'validate filedCd_escritorio_negoci
							Test.ValidadeFileField(p_pathUrlApp, "0;;188;227;Others", "Nm_escritorio_negoci",vNm_escritorio_negoci) 'validate filedNm_escritorio_negoci
							Test.ValidadeFileField(p_pathUrlApp, "0;;228;229;Numeric", "Cd_filial",vCd_filial) 'validate filedCd_filial
							Test.ValidadeFileField(p_pathUrlApp, "0;;230;269;Others", "Nm_filial",vNm_filial) 'validate filedNm_filial
							Test.ValidadeFileField(p_pathUrlApp, "0;;270;272;Alphanumeric", "Cd_regional",vCd_regional) 'validate filedCd_regional
							Test.ValidadeFileField(p_pathUrlApp, "0;;273;312;Others", "Nm_regional",vNm_regional) 'validate filedNm_regional
							Test.ValidadeFileField(p_pathUrlApp, "0;;313;367;Others", "Separador",vSeparador) 'validate filedSeparador
							Test.ValidadeFileField(p_pathUrlApp, "0;;368;392;Others", "Cd_susep_oficial",vCd_susep_oficial) 'validate filedCd_susep_oficial
							Test.ValidadeFileField(p_pathUrlApp, "0;;393;396;Others", "Nm_DDD",vNm_DDD) 'validate filedNm_DDD
							Test.ValidadeFileField(p_pathUrlApp, "0;;397;416;Others", "Nm_fone",vNm_fone) 'validate filedNm_fone
							Test.ValidadeFileField(p_pathUrlApp, "0;;417;457;Others", "Nm_enderecamento_completo",vNm_enderecamento_completo) 'validate filedNm_enderecamento_completo
							Test.ValidadeFileField(p_pathUrlApp, "0;;458;491;Others", "Nm_email",vNm_email) 'validate filedNm_email
							
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
		HandlerError("test_Batch_Cadastro_Corre_Importar_Dados_Corretor.test_Batch_Cadastro_Corre_Importar_Dados_Corretor.Run: " & ex.Message)
		Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
		Return False
            End Try
        End Function

        '****************************************************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vCodigo_Arquivo, vCd_susep, vNm_corretor, vNr_cgc_cpf, vNr_ordem_cgc, vNr_digito_cgc_cpf, vNm_endereco, vNr_endereco, vNm_complemento, vNm_bairro, vNm_cidade, vCd_uf, vCd_cep, vNr_tp_enderecamento, vCd_escritorio_negoci, vNm_escritorio_negoci, vCd_filial, vNm_filial, vCd_regional, vNm_regional, vSeparador, vCd_susep_oficial, vNm_DDD, vNm_fone, vNm_enderecamento_completo, vNm_email, vIsOpenSystem As String

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
                    vCodigo_Arquivo = pc_db.Fieldt("vCodigo_Arquivo")
					vCd_susep = pc_db.Fieldt("vCd_susep")
					vNm_corretor = pc_db.Fieldt("vNm_corretor")
					vNr_cgc_cpf = pc_db.Fieldt("vNr_cgc_cpf")
					vNr_ordem_cgc = pc_db.Fieldt("vNr_ordem_cgc")
					vNr_digito_cgc_cpf = pc_db.Fieldt("vNr_digito_cgc_cpf")
					vNm_endereco = pc_db.Fieldt("vNm_endereco")
					vNr_endereco = pc_db.Fieldt("vNr_endereco")
					vNm_complemento = pc_db.Fieldt("vNm_complemento")
					vNm_bairro = pc_db.Fieldt("vNm_bairro")
					vNm_cidade = pc_db.Fieldt("vNm_cidade")
					vCd_uf = pc_db.Fieldt("vCd_uf")
					vCd_cep = pc_db.Fieldt("vCd_cep")
					vNr_tp_enderecamento = pc_db.Fieldt("vNr_tp_enderecamento")
					vCd_escritorio_negoci = pc_db.Fieldt("vCd_escritorio_negoci")
					vNm_escritorio_negoci = pc_db.Fieldt("vNm_escritorio_negoci")
					vCd_filial = pc_db.Fieldt("vCd_filial")
					vNm_filial = pc_db.Fieldt("vNm_filial")
					vCd_regional = pc_db.Fieldt("vCd_regional")
					vNm_regional = pc_db.Fieldt("vNm_regional")
					vSeparador = pc_db.Fieldt("vSeparador")
					vCd_susep_oficial = pc_db.Fieldt("vCd_susep_oficial")
					vNm_DDD = pc_db.Fieldt("vNm_DDD")
					vNm_fone = pc_db.Fieldt("vNm_fone")
					vNm_enderecamento_completo = pc_db.Fieldt("vNm_enderecamento_completo")
					vNm_email = pc_db.Fieldt("vNm_email")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Batch_Cadastro_Corre_Importar_Dados_Corretor.test_Batch_Cadastro_Corre_Importar_Dados_Corretor.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
