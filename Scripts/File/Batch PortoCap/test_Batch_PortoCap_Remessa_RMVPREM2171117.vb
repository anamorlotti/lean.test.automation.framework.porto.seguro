'**********************************************************************************************************************************************************************
'Create by LeanTest Automation 3.8 in 07/05/2018 18:29:15 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to File
'Environment:..... Automation Project
'Application...... Batch PortoCap
'Functionality:... Remessa RMVPREM2171117
'Master Test:..... No Defined
'TableTest:....... test_Batch_PortoCap_Remessa_RMVPREM2171117
'**********************************************************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Batch_PortoCap_Remessa_RMVPREM2171117
    Public Class test_Batch_PortoCap_Remessa_RMVPREM2171117
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\Files\Saída_Portocap\RMVPREM2171117.TXT"'xml.Read(""urlLocal"") 'Create urlLocal element in XML
								Test.TestLog("Abrir o arquivo Batch PortoCap", "O arquivo  Remessa RMVPREM2171117 deve ser aberto para leitura", "O arquivo Remessa RMVPREM2171117 foi aberto com sucesso", typelog.Passed, , False)
							End If
							Test.ValidadeFileField(p_pathUrlApp, "H;;1;1;Alphabetic", "tipoRegistro",vtipoRegistro) 'validate filedtipoRegistro
							Test.ValidadeFileField(p_pathUrlApp, "H;;2;11;Alphabetic", "Tipo_arquivo",vTipo_arquivo) 'validate filedTipo_arquivo
							Test.ValidadeFileField(p_pathUrlApp, "H;;12;26;Alphabetic", "Nome_empresa",vNome_empresa) 'validate filedNome_empresa
							Test.ValidadeFileField(p_pathUrlApp, "H;;27;41;Numeric", "CGC_empresa",vCGC_empresa) 'validate filedCGC_empresa
							Test.ValidadeFileField(p_pathUrlApp, "H;;42;49;Date", "Data_sorteio",vData_sorteio) 'validate filedData_sorteio
							Test.ValidadeFileField(p_pathUrlApp, "H;;50;55;Numeric", "Versao_arquivo",vVersao_arquivo) 'validate filedVersao_arquivo
							Test.ValidadeFileField(p_pathUrlApp, "H;;56;85;Alphanumeric", "Produto",vProduto) 'validate filedProduto
							Test.ValidadeFileField(p_pathUrlApp, "T;;1;1;Alphabetic", "Tipo_registro_trailler",vTipo_registro_trailler) 'validate filedTipo_registro_trailler
							Test.ValidadeFileField(p_pathUrlApp, "T;;2;7;Numeric", "Quantidade_registros",vQuantidade_registros) 'validate filedQuantidade_registros
							
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
		HandlerError("test_Batch_PortoCap_Remessa_RMVPREM2171117.test_Batch_PortoCap_Remessa_RMVPREM2171117.Run: " & ex.Message)
		Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
		Return False
            End Try
        End Function

        '****************************************************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vtipoRegistro, vTipo_arquivo, vNome_empresa, vCGC_empresa, vData_sorteio, vVersao_arquivo, vProduto, vTipo_registro_trailler, vQuantidade_registros, vIsOpenSystem As String

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
                    vtipoRegistro = pc_db.Fieldt("vtipoRegistro")
					vTipo_arquivo = pc_db.Fieldt("vTipo_arquivo")
					vNome_empresa = pc_db.Fieldt("vNome_empresa")
					vCGC_empresa = pc_db.Fieldt("vCGC_empresa")
					vData_sorteio = pc_db.Fieldt("vData_sorteio")
					vVersao_arquivo = pc_db.Fieldt("vVersao_arquivo")
					vProduto = pc_db.Fieldt("vProduto")
					vTipo_registro_trailler = pc_db.Fieldt("vTipo_registro_trailler")
					vQuantidade_registros = pc_db.Fieldt("vQuantidade_registros")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Batch_PortoCap_Remessa_RMVPREM2171117.test_Batch_PortoCap_Remessa_RMVPREM2171117.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
