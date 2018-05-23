﻿'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 18/05/2018 17:19:29 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to File
'Environment:..... Automation Project
'Application...... XP Paceiro
'Functionality:... Cota
'Master Test:..... No Defined
'TableTest:....... test_XP_Paceiro_Cota
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_XP_Paceiro_Cota
    Public Class test_XP_Paceiro_Cota
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\Files\XP Parceiros\XPCOTA\XPCOTA_2303.TXT"'xml.Read(""urlLocal"") 'Create urlLocal element in XML
								Test.TestLog("Abrir o arquivo XP Paceiro", "O arquivo  Cota deve ser aberto para leitura", "O arquivo Cota foi aberto com sucesso", typelog.Passed, , False)
							End If
							Test.ValidadeFileField(p_pathUrlApp, "0;;1;1;Numeric", "Tipo_de_Registro",vTipo_de_Registro) 'validate filedTipo de Registro
							Test.ValidadeFileField(p_pathUrlApp, "0;;2;9;Date", "Data_atual",vData_atual) 'validate filedData atual
							Test.ValidadeFileField(p_pathUrlApp, "0;;10;29;Alphabetic", "POSICAO_DIARIA",vPOSICAO_DIARIA) 'validate filedPOSICAO DIARIA
							Test.ValidadeFileField(p_pathUrlApp, "0;;30;49;Alphabetic", "COTA",vCOTA) 'validate filedCOTA
							Test.ValidadeFileField(p_pathUrlApp, "0;;50;71;Alphanumeric", "Nome_do_Arquivo",vNome_do_Arquivo) 'validate filedNome do Arquivo
							Test.ValidadeFileField(p_pathUrlApp, "2;;72;72;Numeric", "Tipo_de_Registro_1",vTipo_de_Registro_1) 'validate filedTipo de Registro_1
							Test.ValidadeFileField(p_pathUrlApp, "2;;73;80;Date", "Data_do_Indice",vData_do_Indice) 'validate filedData do Indice
							Test.ValidadeFileField(p_pathUrlApp, "2;;81;85;Alphanumeric", "Codigo_do_fundo",vCodigo_do_fundo) 'validate filedCódigo do fundo
							Test.ValidadeFileField(p_pathUrlApp, "2;;86;101;Numeric", "Valor_da_Cota",vValor_da_Cota) 'validate filedValor da Cota
							Test.ValidadeFileField(p_pathUrlApp, "2;;102;117;Numeric", "Valor_do_Saldo_do_Fundo",vValor_do_Saldo_do_Fundo) 'validate filedValor do Saldo do Fundo
							
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
				HandlerError("test_XP_Paceiro_Cota.test_XP_Paceiro_Cota.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vTipo_de_Registro, vData_atual, vPOSICAO_DIARIA, vCOTA, vNome_do_Arquivo, vTipo_de_Registro_1, vData_do_Indice, vCodigo_do_fundo, vValor_da_Cota, vValor_do_Saldo_do_Fundo, vIsOpenSystem As String

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
                    vTipo_de_Registro = pc_db.Fieldt("vTipo_de_Registro")
					vData_atual = pc_db.Fieldt("vData_atual")
					vPOSICAO_DIARIA = pc_db.Fieldt("vPOSICAO_DIARIA")
					vCOTA = pc_db.Fieldt("vCOTA")
					vNome_do_Arquivo = pc_db.Fieldt("vNome_do_Arquivo")
					vTipo_de_Registro_1 = pc_db.Fieldt("vTipo_de_Registro_1")
					vData_do_Indice = pc_db.Fieldt("vData_do_Indice")
					vCodigo_do_fundo = pc_db.Fieldt("vCodigo_do_fundo")
					vValor_da_Cota = pc_db.Fieldt("vValor_da_Cota")
					vValor_do_Saldo_do_Fundo = pc_db.Fieldt("vValor_do_Saldo_do_Fundo")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_XP_Paceiro_Cota.test_XP_Paceiro_Cota.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
