	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''   Aim                                    :    Aim of the Scirpt  is to run the scripts based on user choice in the test file
''   Reuable  Actions Used   :  
''   Functions  Called            :  FW_Login
''   Test Data  File                 :   FW_UI_0000.xls        
''   Prepared  By                   :   Shobhit Dewan
'' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next
TotstrtTimer=Timer
Public TCID
Public Descr
Public DT
Public TotCasesExecuted
Public TotCasesPassed
Public TotCasesFailed
Public StartTime
Public EndTime
TotCasesExecuted=0
TotCasesPassed=0
TotCasesFailed=0
'Set fso= CreateObject("Scripting.FileSystemObject")
'Set f = fso.GetFolder("C:\Hudson\Jobs\Thor_Trunk\workspace\FW_UI\Scripts\Library Functions\")
'Set fc = f.files
'For Each singlefile in fc
'	If instr (singlefile.name,".lck") <=0 and instr (singlefile.name,".bak")<=0 Then
'		LoadFunctionLibrary (singlefile.name)
'	End If
'Next
'Lod config function is called to set the environment variable Root
call Create_Excel()
Environment.Value("TestCaseID")="Test"
'Call UI_Load_Config("C:\Program Files\Thomson Reuters\Lynx_Editor\Lynx.exe","C:\Program Files\Thomson Reuters\Lynx_Editor\")
'Test Data present in  FW_UI_000.xls  is imported
Config_path = Pathfinder.Locate("Configuration.xml")
Environment.LoadFromFile(Config_path)
Datatable.ImportSheet Environment.Value("Root")&"FW_UI\Test Data\FW_UI_0000.xls",1,2
rowcnt=datatable.getsheet("Action1").GetRowCount
For i = 1 To rowcnt
		ScriptName = DataTable("ScriptName", dtLocalSheet)
		RunTest = DataTable("RunTest", dtLocalSheet)
		Tc_No = Environment.Value("TestName")
		DT = Now
		If (UCase(RunTest) = "YES") Then
				'If Environment.Value("ScriptName")<> ScriptName Then
						TotCasesExecuted=TotCasesExecuted+1
				'End If
				TCID= DataTable("TC_ID", dtLocalSheet)
				Descr= DataTable("Description", dtLocalSheet)
				Params= DataTable("Parameters", dtLocalSheet)
				call Report_Log(TCID,Descr,DT)
				Environment.Value("StartTimer") = Timer
				'' for Load testing of manual publish
				'If TCID="FW_UI_0007" and ScriptName="Verify_Alert_Publish" Then
				'	tempparam=Split(Params,",")
				'	NewParams=tempparam(0)& "," & tempparam(1) & "," & """" & Replace(tempparam(2),"""","") & time() & """"
				'	For j = 1 To 3 Step 1
				'			Execute ScriptName & NewParams
				'	Next
				'else
					Execute ScriptName & Params
				'End If
				''end of editing
				datatable.getsheet("Action1").SetNextRow 
				If DataTable("TC_ID", dtLocalSheet)=TCID Then
				else
					EndTime = Timer
					Datatable.getSheet("Results").setCurrentRow(Environment.Value("TCstrtRw"))
				    Datatable("Time_Taken","Results") = func_ExecutionTime(EndTime-Environment.Value("StartTimer"))
				    Datatable.ExportSheet sTablePath,"Results"
				    Datatable.getSheet("Results").setCurrentRow(Environment.Value("TCendRw")+1)
				    Environment.Value("ScriptName")=ScriptName
				End If 
				datatable.getsheet("Action1").SetPrevRow
		End If
		Datatable.getSheet("Action1").SetNextRow
Next
Call DecorateExcel("C:\Hudson\Jobs\Thor_Trunk\workspace\FW_UI\Test Results\FW_Automation_TestResults"&Date&".xls",TotCasesExecuted,TotstrtTimer)
SystemUtil.CloseProcessByName("Lynx.exe")
msgbox "Completed Execution"
