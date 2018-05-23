'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 21/05/2018 17:13:14 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to DBF
'Environment:..... Automation Project
'Application...... PDL
'Functionality:... CONT.DBF
'Master Test:..... No Defined
'TableTest:....... test_PDL_CONT_DBF
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_PDL_CONT_DBF
    Public Class test_PDL_CONT_DBF
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\Files\PDL\SA_XRE.DBF"'xml.Read(""urlLocal"") 'Create urlLocal element in XML
								Test.TestLog("Abrir o arquivo PDL", "O arquivo  CONT.DBF deve ser aberto para leitura", "O arquivo CONT.DBF foi aberto com sucesso", typelog.Passed, , False)
							End If
							Test.ValidadeDBFField(p_pathUrlApp, ";Numeric;5", "cod_cia",vcod_cia) 'validate filed DBFcod cia
							Test.ValidadeDBFField(p_pathUrlApp, ";Numeric;14", "cnpj_cia",vcnpj_cia) 'validate filed DBFcnpj cia
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;14", "operacao",voperacao) 'validate filed DBFoperacao
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;8", "dt_iniapu",vdt_iniapu) 'validate filed DBFdt iniapu
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;8", "dt_inicont",vdt_inicont) 'validate filed DBFdt inicont
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;8", "dt_fimcont",vdt_fimcont) 'validate filed DBFdt fimcont
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;20", "cpf_partic",vcpf_partic) 'validate filed DBFcpf partic
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;20", "cpf_credit",vcpf_credit) 'validate filed DBFcpf credit
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;60", "nom_titcpf",vnom_titcpf) 'validate filed DBFnom titcpf
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;8", "cd_instit",vcd_instit) 'validate filed DBFcd instit
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;60", "nm_instit",vnm_instit) 'validate filed DBFnm instit
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;20", "cpf_instit",vcpf_instit) 'validate filed DBFcpf instit
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;8", "dt_nascim",vdt_nascim) 'validate filed DBFdt nascim
							Test.ValidadeDBFField(p_pathUrlApp, ";Numeric;3", "idade_part",vidade_part) 'validate filed DBFidade part
							Test.ValidadeDBFField(p_pathUrlApp, ";Numeric;3", "cd_ocup",vcd_ocup) 'validate filed DBFcd ocup
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;60", "nm_ocup",vnm_ocup) 'validate filed DBFnm ocup
							Test.ValidadeDBFField(p_pathUrlApp, ";Numeric;10", "vl_salario",vvl_salario) 'validate filed DBFvl salario
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;20", "pep",vpep) 'validate filed DBFpep
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;60", "nm_cidade",vnm_cidade) 'validate filed DBFnm cidade
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;2", "cd_uf",vcd_uf) 'validate filed DBFcd uf
							Test.ValidadeDBFField(p_pathUrlApp, ";Numeric;2", "cd_produto",vcd_produto) 'validate filed DBFcd produto
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;60", "nm_produto",vnm_produto) 'validate filed DBFnm produto
							Test.ValidadeDBFField(p_pathUrlApp, ";Numeric;4", "cd_plano",vcd_plano) 'validate filed DBFcd plano
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;60", "nm_plano",vnm_plano) 'validate filed DBFnm plano
							Test.ValidadeDBFField(p_pathUrlApp, ";Numeric;2", "origem",vorigem) 'validate filed DBForigem
							Test.ValidadeDBFField(p_pathUrlApp, ";Numeric;12", "proposta",vproposta) 'validate filed DBFproposta
							Test.ValidadeDBFField(p_pathUrlApp, ";Numeric;4", "nr_endosso",vnr_endosso) 'validate filed DBFnr endosso
							Test.ValidadeDBFField(p_pathUrlApp, ";Numeric;10", "vl_movim",vvl_movim) 'validate filed DBFvl movim
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;8", "dt_movim",vdt_movim) 'validate filed DBFdt movim
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;20", "ref_movim",vref_movim) 'validate filed DBFref movim
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;20", "tp_movim",vtp_movim) 'validate filed DBFtp movim
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;20", "formapagto",vformapagto) 'validate filed DBFformapagto
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;10", "nr_susep",vnr_susep) 'validate filed DBFnr susep
							Test.ValidadeDBFField(p_pathUrlApp, ";Char;60", "nm_correto",vnm_correto) 'validate filed DBFnm correto
							
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
				HandlerError("test_PDL_CONT_DBF.test_PDL_CONT_DBF.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vcod_cia, vcnpj_cia, voperacao, vdt_iniapu, vdt_inicont, vdt_fimcont, vcpf_partic, vcpf_credit, vnom_titcpf, vcd_instit, vnm_instit, vcpf_instit, vdt_nascim, vidade_part, vcd_ocup, vnm_ocup, vvl_salario, vpep, vnm_cidade, vcd_uf, vcd_produto, vnm_produto, vcd_plano, vnm_plano, vorigem, vproposta, vnr_endosso, vvl_movim, vdt_movim, vref_movim, vtp_movim, vformapagto, vnr_susep, vnm_correto, vIsOpenSystem As String

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
                    vcod_cia = pc_db.Fieldt("vcod_cia")
					vcnpj_cia = pc_db.Fieldt("vcnpj_cia")
					voperacao = pc_db.Fieldt("voperacao")
					vdt_iniapu = pc_db.Fieldt("vdt_iniapu")
					vdt_inicont = pc_db.Fieldt("vdt_inicont")
					vdt_fimcont = pc_db.Fieldt("vdt_fimcont")
					vcpf_partic = pc_db.Fieldt("vcpf_partic")
					vcpf_credit = pc_db.Fieldt("vcpf_credit")
					vnom_titcpf = pc_db.Fieldt("vnom_titcpf")
					vcd_instit = pc_db.Fieldt("vcd_instit")
					vnm_instit = pc_db.Fieldt("vnm_instit")
					vcpf_instit = pc_db.Fieldt("vcpf_instit")
					vdt_nascim = pc_db.Fieldt("vdt_nascim")
					vidade_part = pc_db.Fieldt("vidade_part")
					vcd_ocup = pc_db.Fieldt("vcd_ocup")
					vnm_ocup = pc_db.Fieldt("vnm_ocup")
					vvl_salario = pc_db.Fieldt("vvl_salario")
					vpep = pc_db.Fieldt("vpep")
					vnm_cidade = pc_db.Fieldt("vnm_cidade")
					vcd_uf = pc_db.Fieldt("vcd_uf")
					vcd_produto = pc_db.Fieldt("vcd_produto")
					vnm_produto = pc_db.Fieldt("vnm_produto")
					vcd_plano = pc_db.Fieldt("vcd_plano")
					vnm_plano = pc_db.Fieldt("vnm_plano")
					vorigem = pc_db.Fieldt("vorigem")
					vproposta = pc_db.Fieldt("vproposta")
					vnr_endosso = pc_db.Fieldt("vnr_endosso")
					vvl_movim = pc_db.Fieldt("vvl_movim")
					vdt_movim = pc_db.Fieldt("vdt_movim")
					vref_movim = pc_db.Fieldt("vref_movim")
					vtp_movim = pc_db.Fieldt("vtp_movim")
					vformapagto = pc_db.Fieldt("vformapagto")
					vnr_susep = pc_db.Fieldt("vnr_susep")
					vnm_correto = pc_db.Fieldt("vnm_correto")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_PDL_CONT_DBF.test_PDL_CONT_DBF.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
