'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 18/05/2018 12:52:07 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to File
'Environment:..... Automation Project
'Application...... Aporte Corretor
'Functionality:... APVD062I02
'Master Test:..... No Defined
'TableTest:....... test_Aporte_Corretor_APVD062I02
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Aporte_Corretor_APVD062I02
    Public Class test_Aporte_Corretor_APVD062I02
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\Files\Aporte Corretor\APVD062I02"'xml.Read(""urlLocal"") 'Create urlLocal element in XML
								Test.TestLog("Abrir o arquivo Aporte Corretor", "O arquivo  APVD062I02 deve ser aberto para leitura", "O arquivo APVD062I02 foi aberto com sucesso", typelog.Passed, , False)
							End If
							Test.ValidadeFileField(p_pathUrlApp, ";;1;2;Others", "Sucursal",vSucursal) 'validate filedSucursal
							Test.ValidadeFileField(p_pathUrlApp, ";;3;6;Numeric", "Ramo",vRamo) 'validate filedRamo
							Test.ValidadeFileField(p_pathUrlApp, ";;7;16;Numeric", "Apolice",vApolice) 'validate filedApolice
							Test.ValidadeFileField(p_pathUrlApp, ";;17;22;Alphanumeric", "Susep",vSusep) 'validate filedSusep
							Test.ValidadeFileField(p_pathUrlApp, ";;23;27;Numeric", "Holding",vHolding) 'validate filedHolding
							Test.ValidadeFileField(p_pathUrlApp, ";;28;32;Numeric", "Codigo_Grupo",vCodigo_Grupo) 'validate filedCodigo Grupo
							Test.ValidadeFileField(p_pathUrlApp, ";;33;37;Numeric", "Holding_Atual",vHolding_Atual) 'validate filedHolding Atual
							Test.ValidadeFileField(p_pathUrlApp, ";;38;42;Numeric", "Codigo_Grupo_Atual",vCodigo_Grupo_Atual) 'validate filedCodigo Grupo Atual
							Test.ValidadeFileField(p_pathUrlApp, ";;43;59;Numeric", "Valor_Aporte",vValor_Aporte) 'validate filedValor Aporte
							Test.ValidadeFileField(p_pathUrlApp, ";;60;61;Others", "Status",vStatus) 'validate filedStatus
							Test.ValidadeFileField(p_pathUrlApp, ";;62;78;Numeric", "Producao_Cobrada_Corretor_RE",vProducao_Cobrada_Corretor_RE) 'validate filedProdução Cobrada Corretor RE
							Test.ValidadeFileField(p_pathUrlApp, ";;79;95;Numeric", "Producao_Cobrada_Corretor_TRANSP",vProducao_Cobrada_Corretor_TRANSP) 'validate filedProdução Cobrada Corretor TRANSP
							Test.ValidadeFileField(p_pathUrlApp, ";;96;112;Numeric", "Producao_Cobrada_Corretor_AUTO",vProducao_Cobrada_Corretor_AUTO) 'validate filedProdução Cobrada Corretor AUTO
							Test.ValidadeFileField(p_pathUrlApp, ";;113;129;Numeric", "Producao_Cobrada_Corretor_SAUDE",vProducao_Cobrada_Corretor_SAUDE) 'validate filedProdução Cobrada Corretor SAUDE
							Test.ValidadeFileField(p_pathUrlApp, ";;130;146;Numeric", "Producao_Cobrada_Corretor_VIDA",vProducao_Cobrada_Corretor_VIDA) 'validate filedProdução Cobrada Corretor VIDA
							Test.ValidadeFileField(p_pathUrlApp, ";;147;163;Numeric", "Producao_Cobrada_Corretor_GARANTIA",vProducao_Cobrada_Corretor_GARANTIA) 'validate filedProdução Cobrada Corretor GARANTIA
							Test.ValidadeFileField(p_pathUrlApp, ";;164;180;Numeric", "Producao_Cobrada_Corretor_FIANCA",vProducao_Cobrada_Corretor_FIANCA) 'validate filedProdução Cobrada Corretor FIANÇA
							Test.ValidadeFileField(p_pathUrlApp, ";;181;197;Numeric", "Producao_Cobrada_Corretor_VAGO1",vProducao_Cobrada_Corretor_VAGO1) 'validate filedProdução Cobrada Corretor_VAGO1
							Test.ValidadeFileField(p_pathUrlApp, ";;198;214;Numeric", "Producao_Cobrada_Corretor_VAGO2",vProducao_Cobrada_Corretor_VAGO2) 'validate filedProdução Cobrada Corretor_VAGO2
							Test.ValidadeFileField(p_pathUrlApp, ";;215;231;Numeric", "Producao_Cobrada_Corretor_VAGO3",vProducao_Cobrada_Corretor_VAGO3) 'validate filedProdução Cobrada Corretor_VAGO3
							Test.ValidadeFileField(p_pathUrlApp, ";;232;237;Numeric", "Percentual_Sinistralidade_Corretor_RE",vPercentual_Sinistralidade_Corretor_RE) 'validate filedPercentual Sinistralidade Corretor RE
							Test.ValidadeFileField(p_pathUrlApp, ";;238;243;Numeric", "Percentual_Sinistralidade_Corretor_TRANSP",vPercentual_Sinistralidade_Corretor_TRANSP) 'validate filedPercentual Sinistralidade Corretor TRANSP
							Test.ValidadeFileField(p_pathUrlApp, ";;244;249;Numeric", "Percentual_Sinistralidade_Corretor_AUTO",vPercentual_Sinistralidade_Corretor_AUTO) 'validate filedPercentual Sinistralidade Corretor AUTO
							Test.ValidadeFileField(p_pathUrlApp, ";;250;255;Numeric", "Percentual_Sinistralidade_Corretor_SAUDE",vPercentual_Sinistralidade_Corretor_SAUDE) 'validate filedPercentual Sinistralidade Corretor SAUDE
							Test.ValidadeFileField(p_pathUrlApp, ";;256;261;", "Percentual_Sinistralidade_Corretor_VIDA",vPercentual_Sinistralidade_Corretor_VIDA) 'validate filedPercentual Sinistralidade Corretor VIDA
							Test.ValidadeFileField(p_pathUrlApp, ";;262;267;Numeric", "Percentual_Sinistralidade_Corretor_GARANTIA",vPercentual_Sinistralidade_Corretor_GARANTIA) 'validate filedPercentual Sinistralidade Corretor GARANTIA
							Test.ValidadeFileField(p_pathUrlApp, ";;268;273;Numeric", "Percentual_Sinistralidade_Corretor_FIANCA",vPercentual_Sinistralidade_Corretor_FIANCA) 'validate filedPercentual Sinistralidade Corretor FIANCA
							Test.ValidadeFileField(p_pathUrlApp, ";;274;279;Numeric", "Percentual_Sinistralidade_Corretor_VAGO1",vPercentual_Sinistralidade_Corretor_VAGO1) 'validate filedPercentual Sinistralidade Corretor VAGO1
							Test.ValidadeFileField(p_pathUrlApp, ";;280;285;Numeric", "Percentual_Sinistralidade_Corretor_VAGO2",vPercentual_Sinistralidade_Corretor_VAGO2) 'validate filedPercentual Sinistralidade Corretor VAGO2
							Test.ValidadeFileField(p_pathUrlApp, ";;286;291;Numeric", "Percentual_Sinistralidade_Corretor_VAGO3",vPercentual_Sinistralidade_Corretor_VAGO3) 'validate filedPercentual Sinistralidade Corretor VAGO3
							Test.ValidadeFileField(p_pathUrlApp, ";;292;297;Numeric", "Percentual_Sinistralidade_Sucursal_RE",vPercentual_Sinistralidade_Sucursal_RE) 'validate filedPercentual Sinistralidade Sucursal RE
							Test.ValidadeFileField(p_pathUrlApp, ";;298;303;Numeric", "Percentual_Sinistralidade_Sucursal_TRANSP",vPercentual_Sinistralidade_Sucursal_TRANSP) 'validate filedPercentual Sinistralidade Sucursal TRANSP
							Test.ValidadeFileField(p_pathUrlApp, ";;304;309;Numeric", "Percentual_Sinistralidade_Sucursal_AUTO",vPercentual_Sinistralidade_Sucursal_AUTO) 'validate filedPercentual Sinistralidade Sucursal AUTO
							Test.ValidadeFileField(p_pathUrlApp, ";;310;315;Numeric", "Percentual_Sinistralidade_Sucursal_SAUDE",vPercentual_Sinistralidade_Sucursal_SAUDE) 'validate filedPercentual Sinistralidade Sucursal SAUDE
							Test.ValidadeFileField(p_pathUrlApp, ";;316;321;Numeric", "Percentual_Sinistralidade_Sucursal_VIDA",vPercentual_Sinistralidade_Sucursal_VIDA) 'validate filedPercentual Sinistralidade Sucursal VIDA
							Test.ValidadeFileField(p_pathUrlApp, ";;322;327;Numeric", "Percentual_Sinistralidade_Sucursal_GARANTIA",vPercentual_Sinistralidade_Sucursal_GARANTIA) 'validate filedPercentual Sinistralidade Sucursal GARANTIA
							Test.ValidadeFileField(p_pathUrlApp, ";;328;333;Numeric", "Percentual_Sinistralidade_Sucursal_FIANCA",vPercentual_Sinistralidade_Sucursal_FIANCA) 'validate filedPercentual Sinistralidade Sucursal FIANÇA
							Test.ValidadeFileField(p_pathUrlApp, ";;334;339;Numeric", "Percentual_Sinistralidade_Sucursal_VAGO1",vPercentual_Sinistralidade_Sucursal_VAGO1) 'validate filedPercentual Sinistralidade Sucursal VAGO1
							Test.ValidadeFileField(p_pathUrlApp, ";;340;345;Numeric", "Percentual_Sinistralidade_Sucursal_VAGO2",vPercentual_Sinistralidade_Sucursal_VAGO2) 'validate filedPercentual Sinistralidade Sucursal VAGO2
							Test.ValidadeFileField(p_pathUrlApp, ";;346;351;Numeric", "Percentual_Sinistralidade_Sucursal_VAGO3",vPercentual_Sinistralidade_Sucursal_VAGO3) 'validate filedPercentual Sinistralidade Sucursal VAGO3
							Test.ValidadeFileField(p_pathUrlApp, ";;352;357;Numeric", "Percentual_Geral_Sinistralidade_Corretor",vPercentual_Geral_Sinistralidade_Corretor) 'validate filedPercentual Geral Sinistralidade Corretor
							Test.ValidadeFileField(p_pathUrlApp, ";;358;363;Numeric", "Percentual_Sinistralidade_Ponderada",vPercentual_Sinistralidade_Ponderada) 'validate filedPercentual Sinistralidade Ponderada
							
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
				HandlerError("test_Aporte_Corretor_APVD062I02.test_Aporte_Corretor_APVD062I02.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vSucursal, vRamo, vApolice, vSusep, vHolding, vCodigo_Grupo, vHolding_Atual, vCodigo_Grupo_Atual, vValor_Aporte, vStatus, vProducao_Cobrada_Corretor_RE, vProducao_Cobrada_Corretor_TRANSP, vProducao_Cobrada_Corretor_AUTO, vProducao_Cobrada_Corretor_SAUDE, vProducao_Cobrada_Corretor_VIDA, vProducao_Cobrada_Corretor_GARANTIA, vProducao_Cobrada_Corretor_FIANCA, vProducao_Cobrada_Corretor_VAGO1, vProducao_Cobrada_Corretor_VAGO2, vProducao_Cobrada_Corretor_VAGO3, vPercentual_Sinistralidade_Corretor_RE, vPercentual_Sinistralidade_Corretor_TRANSP, vPercentual_Sinistralidade_Corretor_AUTO, vPercentual_Sinistralidade_Corretor_SAUDE, vPercentual_Sinistralidade_Corretor_VIDA, vPercentual_Sinistralidade_Corretor_GARANTIA, vPercentual_Sinistralidade_Corretor_FIANCA, vPercentual_Sinistralidade_Corretor_VAGO1, vPercentual_Sinistralidade_Corretor_VAGO2, vPercentual_Sinistralidade_Corretor_VAGO3, vPercentual_Sinistralidade_Sucursal_RE, vPercentual_Sinistralidade_Sucursal_TRANSP, vPercentual_Sinistralidade_Sucursal_AUTO, vPercentual_Sinistralidade_Sucursal_SAUDE, vPercentual_Sinistralidade_Sucursal_VIDA, vPercentual_Sinistralidade_Sucursal_GARANTIA, vPercentual_Sinistralidade_Sucursal_FIANCA, vPercentual_Sinistralidade_Sucursal_VAGO1, vPercentual_Sinistralidade_Sucursal_VAGO2, vPercentual_Sinistralidade_Sucursal_VAGO3, vPercentual_Geral_Sinistralidade_Corretor, vPercentual_Sinistralidade_Ponderada, vIsOpenSystem As String

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
                    vSucursal = pc_db.Fieldt("vSucursal")
					vRamo = pc_db.Fieldt("vRamo")
					vApolice = pc_db.Fieldt("vApolice")
					vSusep = pc_db.Fieldt("vSusep")
					vHolding = pc_db.Fieldt("vHolding")
					vCodigo_Grupo = pc_db.Fieldt("vCodigo_Grupo")
					vHolding_Atual = pc_db.Fieldt("vHolding_Atual")
					vCodigo_Grupo_Atual = pc_db.Fieldt("vCodigo_Grupo_Atual")
					vValor_Aporte = pc_db.Fieldt("vValor_Aporte")
					vStatus = pc_db.Fieldt("vStatus")
					vProducao_Cobrada_Corretor_RE = pc_db.Fieldt("vProducao_Cobrada_Corretor_RE")
					vProducao_Cobrada_Corretor_TRANSP = pc_db.Fieldt("vProducao_Cobrada_Corretor_TRANSP")
					vProducao_Cobrada_Corretor_AUTO = pc_db.Fieldt("vProducao_Cobrada_Corretor_AUTO")
					vProducao_Cobrada_Corretor_SAUDE = pc_db.Fieldt("vProducao_Cobrada_Corretor_SAUDE")
					vProducao_Cobrada_Corretor_VIDA = pc_db.Fieldt("vProducao_Cobrada_Corretor_VIDA")
					vProducao_Cobrada_Corretor_GARANTIA = pc_db.Fieldt("vProducao_Cobrada_Corretor_GARANTIA")
					vProducao_Cobrada_Corretor_FIANCA = pc_db.Fieldt("vProducao_Cobrada_Corretor_FIANCA")
					vProducao_Cobrada_Corretor_VAGO1 = pc_db.Fieldt("vProducao_Cobrada_Corretor_VAGO1")
					vProducao_Cobrada_Corretor_VAGO2 = pc_db.Fieldt("vProducao_Cobrada_Corretor_VAGO2")
					vProducao_Cobrada_Corretor_VAGO3 = pc_db.Fieldt("vProducao_Cobrada_Corretor_VAGO3")
					vPercentual_Sinistralidade_Corretor_RE = pc_db.Fieldt("vPercentual_Sinistralidade_Corretor_RE")
					vPercentual_Sinistralidade_Corretor_TRANSP = pc_db.Fieldt("vPercentual_Sinistralidade_Corretor_TRANSP")
					vPercentual_Sinistralidade_Corretor_AUTO = pc_db.Fieldt("vPercentual_Sinistralidade_Corretor_AUTO")
					vPercentual_Sinistralidade_Corretor_SAUDE = pc_db.Fieldt("vPercentual_Sinistralidade_Corretor_SAUDE")
					vPercentual_Sinistralidade_Corretor_VIDA = pc_db.Fieldt("vPercentual_Sinistralidade_Corretor_VIDA")
					vPercentual_Sinistralidade_Corretor_GARANTIA = pc_db.Fieldt("vPercentual_Sinistralidade_Corretor_GARANTIA")
					vPercentual_Sinistralidade_Corretor_FIANCA = pc_db.Fieldt("vPercentual_Sinistralidade_Corretor_FIANCA")
					vPercentual_Sinistralidade_Corretor_VAGO1 = pc_db.Fieldt("vPercentual_Sinistralidade_Corretor_VAGO1")
					vPercentual_Sinistralidade_Corretor_VAGO2 = pc_db.Fieldt("vPercentual_Sinistralidade_Corretor_VAGO2")
					vPercentual_Sinistralidade_Corretor_VAGO3 = pc_db.Fieldt("vPercentual_Sinistralidade_Corretor_VAGO3")
					vPercentual_Sinistralidade_Sucursal_RE = pc_db.Fieldt("vPercentual_Sinistralidade_Sucursal_RE")
					vPercentual_Sinistralidade_Sucursal_TRANSP = pc_db.Fieldt("vPercentual_Sinistralidade_Sucursal_TRANSP")
					vPercentual_Sinistralidade_Sucursal_AUTO = pc_db.Fieldt("vPercentual_Sinistralidade_Sucursal_AUTO")
					vPercentual_Sinistralidade_Sucursal_SAUDE = pc_db.Fieldt("vPercentual_Sinistralidade_Sucursal_SAUDE")
					vPercentual_Sinistralidade_Sucursal_VIDA = pc_db.Fieldt("vPercentual_Sinistralidade_Sucursal_VIDA")
					vPercentual_Sinistralidade_Sucursal_GARANTIA = pc_db.Fieldt("vPercentual_Sinistralidade_Sucursal_GARANTIA")
					vPercentual_Sinistralidade_Sucursal_FIANCA = pc_db.Fieldt("vPercentual_Sinistralidade_Sucursal_FIANCA")
					vPercentual_Sinistralidade_Sucursal_VAGO1 = pc_db.Fieldt("vPercentual_Sinistralidade_Sucursal_VAGO1")
					vPercentual_Sinistralidade_Sucursal_VAGO2 = pc_db.Fieldt("vPercentual_Sinistralidade_Sucursal_VAGO2")
					vPercentual_Sinistralidade_Sucursal_VAGO3 = pc_db.Fieldt("vPercentual_Sinistralidade_Sucursal_VAGO3")
					vPercentual_Geral_Sinistralidade_Corretor = pc_db.Fieldt("vPercentual_Geral_Sinistralidade_Corretor")
					vPercentual_Sinistralidade_Ponderada = pc_db.Fieldt("vPercentual_Sinistralidade_Ponderada")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Aporte_Corretor_APVD062I02.test_Aporte_Corretor_APVD062I02.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
