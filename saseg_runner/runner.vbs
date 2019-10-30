Option Explicit

Dim app         ' As SASEGuide.Application
Call dowork
'shut down the app
If not (app Is Nothing) Then
    app.Quit
    Set app = Nothing
End If

Sub dowork()
    On Error Resume Next
    '----
    ' Start up Enterprise Guide using the project name
    '----
    Dim prjName     ' As String
    Dim prjObject   ' As SASEGuide.Project
    Dim containerName     ' As String
    Dim containerObject   ' As SASEGuide.Container
    Dim containerColl     ' As SASEGuide.ContainerCollection
    prjName = "test.egp" ' Project Name
    containerName = "Process Flow" ' Container Name
    Set app = CreateObject("SASEGObjectModel.Application.7.1")
    If Checkerror("CreateObject") = True Then
        Exit Sub
    End If
    Set prjObject = app.Open(prjName,"")
    If Checkerror("App.Open") = True Then
        Exit Sub
    End If
    '-----
    'Get The Container Collection and Object
    '-----   
    Set containerColl = prjObject.ContainerCollection
    If Checkerror("Project.ContainerCollection") = True Then
        Exit Sub
    End If
    Dim i       ' As Long
    Dim count   ' As Long
    count = containerColl.count
    For i = 0 To count - 1
        Set containerObject = containerColl.Item(i)
        If Checkerror("ContainerCollection.Item") = True Then
            Exit Sub
        End If
        If (containerObject.Name = containerName) Then
            Exit For
        Else
            Set containerObject = Nothing
        End If
    Next
    If not (containerObject Is Nothing) Then
        '----
        ' Run the Container
        '----
        containerObject.Run
        If Checkerror("Container.Run") = True Then
            Exit Sub
        End If              
    End If
    '-----
    ' Save the new project
    '-----

    prjObject.Save
    If Checkerror("Project.Save") = True Then
        Exit Sub
    End If
    '-----
    ' Close the project
    '-----

    prjObject.Close
    If Checkerror("Project.Close") = True Then
        Exit Sub
    End If
End Sub

Function Checkerror(fnName)
    Checkerror = False
    Dim strmsg      ' As String
    Dim errNum      ' As Long
    If Err.Number <> 0 Then
        strmsg = "Error #" & Hex(Err.Number) & vbCrLf & "In Function " & fnName & vbCrLf & Err.Description
        MsgBox strmsg  'Uncomment this line if you want to be notified via MessageBox of Errors in the script.
        Checkerror = True
    End If
End Function