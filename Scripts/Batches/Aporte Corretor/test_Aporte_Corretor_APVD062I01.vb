'*********************************************************************************************************************************
'Create by LeanTest Automation 3.6 in 18/05/2018 11:35:04 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to File
'Environment:..... Automation Project
'Application...... Aporte Corretor
'Functionality:... APVD062I01
'Master Test:..... No Defined
'TableTest:....... test_Aporte_Corretor_APVD062I01
'*********************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_Aporte_Corretor_APVD062I01
    Public Class test_Aporte_Corretor_APVD062I01
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                    	Try
                    		If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "C:\LeanTestAutomation\Scripts\lean.test.automation.framework.porto.seguro\Files\Aporte Corretor\APVD062I01"'xml.Read(""urlLocal"") 'Create urlLocal element in XML
								Test.TestLog("Abrir o arquivo Aporte Corretor", "O arquivo  APVD062I01 deve ser aberto para leitura", "O arquivo APVD062I01 foi aberto com sucesso", typelog.Passed, , False)
							End If
                            Test.ValidadeFileField(p_pathUrlApp, ";;1;2;Others", "Sucursal",vSucursal) 'validate filedSucursal
							Test.ValidadeFileField(p_pathUrlApp, ";;3;6;Numeric", "Ramo",vRamo) 'validate filedRamo
							Test.ValidadeFileField(p_pathUrlApp, ";;7;16;Numeric", "Apolice",vApolice) 'validate filedApolice
							Test.ValidadeFileField(p_pathUrlApp, ";;17;22;Alphanumeric", "Susep",vSusep) 'validate filedSusep
							Test.ValidadeFileField(p_pathUrlApp, ";;23;27;Numeric", "Holding",vHolding) 'validate filedHolding
							Test.ValidadeFileField(p_pathUrlApp, ";;28;32;Numeric", "Codigo_Grupo",vCodigo_Grupo) 'validate filedCodigo Grupo
							Test.ValidadeFileField(p_pathUrlApp, ";;33;40;Numeric", "Perc_Participacao_Aporte",vPerc_Participacao_Aporte) 'validate filedPerc. Participação Aporte
							Test.ValidadeFileField(p_pathUrlApp, ";;41;41;Others", "Flag",vFlag) 'validate filedFlag
							Test.ValidadeFileField(p_pathUrlApp, ";YYYY-MM;42;48;Others", "Mes_Referencia",vMes_Referencia) 'validate filedMês Referencia
							
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
				HandlerError("test_Aporte_Corretor_APVD062I01.test_Aporte_Corretor_APVD062I01.Run: " & ex.Message)
                Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
                Return False
            End Try
        End Function

        '*********************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vSucursal, vRamo, vApolice, vSusep, vHolding, vCodigo_Grupo, vPerc_Participacao_Aporte, vFlag, vMes_Referencia, vIsOpenSystem As String

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
					vPerc_Participacao_Aporte = pc_db.Fieldt("vPerc_Participacao_Aporte")
					vFlag = pc_db.Fieldt("vFlag")
					vMes_Referencia = pc_db.Fieldt("vMes_Referencia")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
										
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_Aporte_Corretor_APVD062I01.test_Aporte_Corretor_APVD062I01.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
