' ------------------------------------------------
' ExtractCodeAndLog.vbs
' Uses SAS Enterprise Guide automation to read an EGP file
' and export all SAS programs and logs to subfolders
' There is a new subfolder created for each process flow
' within the project
'
' This script uses the Code.Text method, which will NOT include
' other "wrapper" code around each program that 
' SAS Enterprise Guide might have inserted
'
' USAGE:
'   cscript.exe ExtractCodeAndLog.vbs <path-to-EGP-file>
' EXAMPLE:
'   cscript.exe ExtractCodeAndLog.vbs c:\projects\DonorsChoose.egp
'
' NOTE: use proper version of CSCRIPT.exe for your SAS Enterprise Guide
' version.  For 32-bit EG on 64-bit Windows, use 
'      c:\Windows\SysWOW64\cscript.exe 
' ------------------------------------------------
' force declaration of variables in VB Script
Option Explicit
Dim Application
Dim Project

' Change if running a different version of EG
Dim egVersion 

' enumeration of project item types
Const egLog = 0  
Const egCode = 1  
Const egData = 2  
Const egQuery = 3  
Const egContainer = 4  
Const egDocBuilder = 5  
Const egNote = 6  
Const egResult = 7  
Const egTask = 8  
Const egTaskCode = 9  
Const egProjectParameter = 10  
Const egOutputData = 11  
Const egStoredProcess = 12  
Const egStoredProcessParameter = 13  
Const egPublishAction = 14  
Const egCube = 15  
Const egReport = 18  
Const egReportSnapshot = 19  
Const egOrderedList = 20  
Const egSchedule = 21  
Const egLink = 22  
Const egFile = 23  
Const egIntrNetApp = 24  
Const egInformationMap = 25  

' Simple error check - looking for a project file 
If WScript.Arguments.Count = 0 Then
  WScript.Echo "ERROR: Expecting the full path name of a project file"
  WScript.Quit -1
End If

egVersion = "SASEGObjectModel.Application." & WScript.Arguments.Item(1)

' Create a new SAS Enterprise Guide automation session
On Error Resume Next
Set Application = WScript.CreateObject(egVersion)
WScript.Echo "Opening project: " & WScript.Arguments.Item(0)

' Open the EGP file with the Application
Set Project = Application.Open(WScript.Arguments.Item(0),"")
If Err.Number <> 0 Then
  WScript.Echo "ERROR: Unable to open " _
    & WScript.Arguments.Item(0) & " as a project file"
  WScript.Quit -1
End If

' Create a subfolder in the current directory
' for these programs
Dim programsFolder
programsFolder = getWorkingDirectory() _
   & "\" & Project.Name & "_CodeAndLogs"
Wscript.Echo "Target folder for SAS programs is "  & programsFolder
createFolder(programsFolder)
Dim item 
Dim flow

' Navigate the process flows in the Project object
For Each flow In Project.ContainerCollection
  ' ProcessFlow is ContainerType of 0
  If flow.ContainerType = 0 Then
    Dim tableOfContents
    Dim progNumber
	progNumber = 1
    WScript.Echo "Process Flow: " & flow.Name
	Dim flowFolder
    flowFolder = programsFolder & "\" & flow.Name	
	createFolder(flowFolder)
	' Navigate the items in each process flow
	tableOfContents = "Contents of flow: " & flow.Name & vbNewLine & "----------------" 
	For Each item in flow.Items
	  tableOfContents = tableOfContents & vbNewLine & item.Name
	  progNumber = progNumber + 1
	  Select Case item.Type
	  Case egFile 
	    WScript.Echo "  " & item.Name & ", File Name: " & item.FileName
		tableOfContents = tableOfContents & " (External file)" 
	  Case egCode 
	    WScript.Echo "  " & item.Name & ", SAS Program"  
		tableOfContents = tableOfContents & " (SAS Program)" 
		WScript.Echo "    saving program content to: " _
		   & flowFolder & "\" & progNumber & "_" & item.Name & ".sas " _
		   & vbNewLine & "   (" & Len(item.Text) & " characters)"
		' Save just the text of this program to an external file
		saveTextToFile flowFolder & "\" & progNumber & "_" & item.Name & ".sas", item.Text   
		If (Not item.Log Is Nothing) Then
		  ' if there is a log, save the text of the log to an external file
		  saveTextToFile flowFolder & "\" & progNumber & "_" & item.Name & ".log", item.Log.Text   
		End If
		
		If Err.Number <> 0 Then
		  WScript.Echo "     ERROR: There was an error while saving " _
		    & flowFolder & "\" & item.Name & ".sas: " & vbNewLine & Err.Description
		  Err.Clear
		End If
	  ' when processing the DATA items, iterate through
	  ' the subtasks that exist, if any
	  Case egData
	    WScript.Echo "  " & item.Name & ", Data: " & item.FileName
		tableOfContents = tableOfContents & " (Data set)" 
		Dim task
		For Each task in item.Tasks
		  tableOfContents = tableOfContents & vbNewLine & "   " & task.Name & " (Task)"
		  If (Not task.TaskCode is Nothing) Then
		    saveTextToFile flowFolder & "\" & progNumber & "_" & task.Name & ".sas", task.TaskCode.Text   
		  End If
		  If (Not task.Log is Nothing) Then
		    saveTextToFile flowFolder & "\" & progNumber & "_" & task.Name & ".log", task.Log.Text   
		  End If
		Next
		
	  Case egTask
	    WScript.Echo "  " & item.Name & ", Task"
		tableOfContents = tableOfContents & " (Task)" 	
		If (Not item.TaskCode is Nothing) Then
		  saveTextToFile flowFolder & "\" & progNumber & "_" & item.Name & ".sas", item.TaskCode.Text   
		End If
		If (Not item.Log is Nothing) Then
		  saveTextToFile flowFolder & "\" & progNumber & "_" & item.Name & ".log", item.Log.Text   
		End If		
		
	  Case egQuery
	    WScript.Echo "  " & item.Name & ", Query"
		tableOfContents = tableOfContents & " (Query)" 		
		If (Not item.TaskCode is Nothing) Then
		  saveTextToFile flowFolder & "\" & progNumber & "_" & item.Name & ".sas", item.TaskCode.Text   
		End If
		If (Not item.Log is Nothing) Then
		  saveTextToFile flowFolder & "\" & progNumber & "_" & item.Name & ".log", item.Log.Text   
		End If				
	  End Select
	Next
	saveTextToFile flowFolder & "\" & flow.Name & ".txt", tableOfContents
  End If
Next

' Close the project
Project.Close
' Quit/close the Application object
Application.Quit

' --- Helper functions ----------------

' Save a block of text (code or log) to text file
Function saveTextToFile(filename, text)
  Dim objFS
  Dim objOutFile
  Set objFS = CreateObject("Scripting.FileSystemObject")
  Set objOutFile = objFS.CreateTextFile(filename,True)
  objOutFile.Write(text)
  objOutFile.Close
End Function

' function to get the current working directory
Function getWorkingDirectory()
	Dim objShell
	Set objShell = CreateObject("Wscript.Shell")
	getWorkingDirectory = objShell.CurrentDirectory
End Function

' function to create a new subfolder, if it doesn't yet exist
Function createFolder(folderName)
  Dim objFSO
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  If objFSO.FolderExists(folderName) = False Then
	objFSO.CreateFolder folderName
  End If
End Function
