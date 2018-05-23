'**********************************************************************************************************************************************************************
'Create by LeanTest Automation 3.8 in 27/04/2018 12:27:09 (By LeanTest Automation Test) 
'User:............ Admin
'Domain:.......... LeanTest Execution Automation to API Rest
'Environment:..... Automation Project
'Application...... OSB Atlas Buscar Cot
'Functionality:... OSB BuscarCota
'Master Test:..... No Defined
'TableTest:....... test_OSB_Atlas_Buscar_Cot_OSB_BuscarCota
'**********************************************************************************************************************************************************************
Imports Lean.Test.Automation.Framework.LibraryGlobal.LibGlobal
Namespace test_OSB_Atlas_Buscar_Cot_OSB_BuscarCota
    Public Class test_OSB_Atlas_Buscar_Cot_OSB_BuscarCota
        Public Sub New()
        End Sub
        Public Function Run() As Boolean
            Try
                If StartTest() Then
                    Do While p_CountTest <> 0
                        Try
                            If CBool(vIsOpenSystem) Then 'if true open system else load new test without open system
								p_pathUrlApp = "http://nt1156:803/atlassatwebapi/api/CarteiraCota/BuscarCotaRequest"'xml.Read("urlLocal") 'Create urlLocal element in XML
								If Not Test.StartTests Then Return False
							End If
							vJsonParameters = Test.GetValueOutput(p_IDTest, p_TableTest, "vJsonParameters") 'get JsonParameters OSB Atlas Buscar Cot
							vJsonRequest = Test.GetValueOutput(p_IDTest, p_TableTest, "vJsonRequest") 'get JsonRequest OSB Atlas Buscar Cot
							vJsonRequest = Test.SetValueJson("vIdCarteira", vIdCarteira,"", vJsonRequest) 'type value in element jsonIdCarteira
							vJsonRequest = Test.SetValueJson("vData", vData,"", vJsonRequest) 'type value in element jsonData
							Test.TestLog("Consumir o serviço OSB BuscarCota", "O serviço OSB BuscarCota deve ser consumido com sucesso usando os valores " & vbCrLf & vJsonRequest, "O serviço OSB BuscarCota foi consumido com sucesso", typelog.Passed, , False)
							p_APIIRestLeanTest = Test.CallAPIRest(p_pathUrlApp, vJsonRequest, vJsonParameters, p_TypeExecution) 'CallAPIRest OSB Atlas Buscar Cot
							If p_APIIRestLeanTest.StatusCode = 200 Then
								Test.TestLog("Verificar o retorno do serviço OSB BuscarCota", "O serviço OSB BuscarCota retornou com sucesso usando os valores " & vbCrLf & p_APIIRestLeanTest.InnerJson, "O serviço OSB BuscarCota foi consumido com sucesso", typelog.Passed, , False)
							Else
							Test.TestLog("Verificar o retorno do serviço OSB BuscarCota", "O serviço OSB BuscarCota retornou a mensagem de erro: " & vbCrLf & p_APIIRestLeanTest.InnerJson, "O serviço OSB BuscarCota não foi consumido com sucesso", typelog.Failed, , False)
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
		HandlerError("test_OSB_Atlas_Buscar_Cot_OSB_BuscarCota.test_OSB_Atlas_Buscar_Cot_OSB_BuscarCota.Run: " & ex.Message)
		Test.TestLog("Execução do teste", "Teste executado com sucesso", "Teste executado com falha! Message: " & p_errorDescription, typelog.Failed)
		Return False
            End Try
        End Function

        '****************************************************************************************************************************************************************
        'STARTTEST
        Public Shared p_ExpectedResult, p_CheckPoint1 As String
        Public Shared vJsonParameters,vJsonRequest, vIdCarteira, vData, vIsOpenSystem As String

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
                    vJsonParameters = pc_db.Fieldt("vJsonParameters")
					vJsonRequest = pc_db.Fieldt("vJsonRequest")
					vIdCarteira = pc_db.Fieldt("vIdCarteira")
					vData = pc_db.Fieldt("vData")
					vIsOpenSystem = pc_db.Fieldt("vIsOpenSystem")
					
                    
                    pc_db.StartExecution(p_TableTest, p_IDTest)
                    If p_PublishQC Then CreateStructureQC()
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                HandlerError("test_OSB_Atlas_Buscar_Cot_OSB_BuscarCota.test_OSB_Atlas_Buscar_Cot_OSB_BuscarCota.StartTest" & ex.StackTrace & " - " & ex.Message)
                Return False
            End Try
        End Function
    End Class
End Namespace
