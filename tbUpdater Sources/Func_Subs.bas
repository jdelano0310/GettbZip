Attribute VB_Name = "Func_Subs"

' add your procedures here
Public tbUpdaterSettings As clsSettings
Public fso As FileSystemObject
Public chgLogs As New colChangeLogItems

Public Function askForFolder(dlgCaption As String, Optional currentFolder As String) As String
        
    ' open the select folder form and return the selection (if any)
    Dim frmSelFolder As New frmSelectFolder
    
    With frmSelFolder
        If fso.FolderExists(currentFolder) Then .selectedFolder = currentFolder
        .Caption = dlgCaption
        .Show(vbModal)
    End With
    
    Return frmSelFolder.selectedFolder
        
End Function

Public Function FillLogHistoryGrid(Optional ViewDate As String = "") As Boolean
    
    ' open the log text file and display it in the flexgrid
    Dim logContents As New colHistoryLogItems
    Dim logItem As clsHistoryLogItem
    Dim itemColor As Long
    Dim colNum As Integer
    
    If Not logContents.LoadLog Then
        MsgBox("There was an issue reading the log file", vbExclamation, "View log")
        FillLogHistoryGrid = False
        Exit Function
    End If
    
    frmViewLog.flgLog.Rows = 1
    For Each logItem In logContents
        With frmViewLog.flgLog
            If Len(ViewDate) = 0 Or logItem.LogDate = ViewDate Or ViewDate = "Show All" Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
            
                If logItem.LogCLI.tBVersion = 0 Then
                    .TextMatrix(.Row, 0) = logItem.LogDateTime
                    .TextMatrix(.Row, 3) = logItem.LogMessage
                Else
                    .TextMatrix(.Row, 0) = logItem.LogDateTime
                    .TextMatrix(.Row, 1) = logItem.LogCLI.tBVersion
                    .TextMatrix(.Row, 2) = logItem.LogCLI.Type
                    .TextMatrix(.Row, 3) = logItem.LogCLI.Notes
                End If
                
                ' color the change log of the tB version that was installed
                ' during the logging process
                itemColor = ChangeLogItemColor(logItem.LogCLI.Type)
                For colNum = 0 To .Cols - 1
                    .Col = colNum
                    .CellForeColor = itemColor
                Next
            End If
        End With
        DoEvents()
    Next
    
    If Len(ViewDate) = 0 Then
        ' fill the dropdown with the unique dates from the log file
        ' just during the first time through this
        Dim logDate As String
        
        With frmViewLog.cboLogDate
            .Clear()
            .AddItem("Show All")
            For Each logDate In logContents.HistoryLogDates
                .AddItem(logDate)
            Next
            .ListIndex = 0
        End With
    End If
    
    FillLogHistoryGrid = True
End Function

Public Sub FilltBChangeLog(Optional fortBVersion As String = "")
    
    ' open the log text file and display it in the flexgrid
    Dim clItem As clsChangeLogItem
    Dim itemColor As Long
    Dim colNum As Integer
    
    If chgLogs.tBVersionGap > 1 Then
        Dim gridTitleCaption As String
        
        ' make the caption seem more correct
        If chgLogs.tBVersionGap = 2 Then
            gridTitleCaption = CStr(chgLogs.InstalledtBVersion + 1) & " and " & chgLogs.LatestVersion
        Else
            gridTitleCaption = CStr(chgLogs.InstalledtBVersion + 1) & " thru " & chgLogs.LatestVersion
        End If
        
        Form1.lblChangeLogTitle.Caption = "Change logs for " & gridTitleCaption
    ElseIf chgLogs.tBVersionGap = 1 Then
        Form1.lblChangeLogTitle.Caption = "Change Log for " & chgLogs.LatestVersion
    Else
        Form1.lblChangeLogTitle.Caption = "Change Log"
    End If
    
    With Form1.flgLog
        .Rows = 1
        For Each clItem In chgLogs
            
                If Len(fortBVersion) = 0 Or clItem.tBVersion = Val(fortBVersion) Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                
                    .TextMatrix(.Row, 0) = clItem.tBVersion
                    .TextMatrix(.Row, 1) = clItem.Type
                    .TextMatrix(.Row, 2) = clItem.Notes
                    
                    ' color the change log of the tB version that was installed
                    ' during the logging process
                    itemColor = ChangeLogItemColor(clItem.Type)
                    For colNum = 0 To .Cols - 1
                        .Col = colNum
                        .CellForeColor = itemColor
                    Next
                End If
            
            DoEvents()
        Next
    End With
End Sub

Public Function ChangeLogItemColor(changeLogType As String) As Long
    Dim itemColor As Long
    
    Select Case UCase(Trim(changeLogType))
        Case "IMPORTANT"
            itemColor = vbBlue
        Case "KNOWN ISSUE"
            itemColor = vbBlack
        Case "TIP"
            itemColor = RGB(22, 83, 126) ' blueish
        Case "WARNING"
            itemColor = RGB(153, 0, 0)   ' dark red
        Case "FIXED"
            itemColor = RGB(56, 118, 29) ' green
        Case "ADDED"
            itemColor = RGB(75, 0, 130)  ' indigo
        Case "UPDATED"
            itemColor = RGB(0, 128, 0)   ' other green 
        Case Else
            itemColor = vbBlack
    End Select
    
    Return itemColor
    
End Function

Public Function GetCurrentTBVersion(tBFolder As String) As String
        
    ' attempt to find the version number of twinBasic in use
    Dim fileWithVersionInfo As String = tBFolder & "ide\build.js"
    Dim versionIndicator As String = "BETA"
    Dim fileContents As String
    Dim tempString As String
    
    If Not fso.FileExists(fileWithVersionInfo) Then
        GetCurrentTBVersion = "Version file," & fileWithVersionInfo & ", is missing"
        Exit Function
    End If
        
    ' open the file designated as the one with the version number
    fileContents = fsoFileRead(fileWithVersionInfo)
    
    ' parse the text for the version number
    tempString = Mid(fileContents, InStr(fileContents, versionIndicator))
    GetCurrentTBVersion = Mid(tempString, Len(versionIndicator) + 1, 4)
    
    chgLogs.InstalledtBVersion = GetCurrentTBVersion
    
End Function

Public Function fsoFileRead(filePath As String) As String
    
    If Not fso.FileExists(filePath) Then Return "Failed fsoFileRead"
    
    Dim fso As New Scripting.FileSystemObject
        Dim fileToRead As TextStream
        
        Set fileToRead = fso.OpenTextFile(filePath, ForReading)
            fsoFileRead = fileToRead.ReadAll()
        fileToRead.Close()
    Set fso = Nothing
    
End Function

Public Sub LoadSettingsIntoForm()
    
    ' set the controls on the form with their settings values
    With Form1
        .txtDownloadTo.Text = tbUpdaterSettings.DownloadFolder
        .txttBLocation.Text = tbUpdaterSettings.twinBASICFolder
        Select Case tbUpdaterSettings.PostDownloadAction
            Case 1
                .optOpenFolder.Value = True
            Case 2
                .optOpenZip.Value = True
            Case 3
                .optInstallTB.Value = True
        End Select
        .chkLookForUpdateOnLaunch.Value = tbUpdaterSettings.CheckForNewVersionOnLoad
        .chkStarttwinBASIC.Value = tbUpdaterSettings.StarttwinBASICAfterUpdate
        .chkLog.Value = tbUpdaterSettings.LogActivity
        .chkSaveSettings.Value = tbUpdaterSettings.SaveSettingsOnExit
    End With
    
End Sub

Public Sub WriteToLogFile()
        
    ' write the contents of the displayed log to the log history file
    Dim logFile As TextStream
    Dim logFileName As String = App.Path & "\log.txt"
    Dim logIndex As Integer
    Dim tbVersionInstalled As Boolean = False
    
    Dim lb As ListBox = Form1.lbStatus
    Dim grd As VBFlexGrid = Form1.flgLog
    
    Set logFile = fso.OpenTextFile(logFileName, ForAppending, True)
        For logIndex = 0 To lb.ListCount - 1
            logFile.WriteLine(lb.List(logIndex))
            If Not tbVersionInstalled Then tbVersionInstalled = InStr(lb.List(logIndex), "Post download") > 1 ' if the user at least downloaded the zip
        Next logIndex
        
        ' write the change log(s) for the version downloaded, plus the previous versions inbetween 
        ' the installed and the latest available installed
        If tbVersionInstalled Then
            For logIndex = 1 To grd.Rows - 1
                logFile.WriteLine(Format(Now, "MM/dd/yy hh:mm:ss AM/PM: ") & grd.TextMatrix(logIndex, 0) & " - " & grd.TextMatrix(logIndex, 1) & " - " & grd.TextMatrix(logIndex, 2))
            Next logIndex
        End If
    logFile.Close()
    Set logFile = Nothing
        
End Sub

Public Sub LogGridClick(logGrid As VBFlexGrid)
    
    ' override the forecolor for the selected row with the color for the type of record
    Dim selectCLType As String
    Dim typeColNum As Integer
    
    ' the girds that use this have the type col in different places, find the proper col
    For typeColNum = 0 To logGrid.Cols - 1
        If logGrid.TextMatrix(0, typeColNum) = "Type" Then Exit For
    Next
    
    selectCLType = logGrid.TextMatrix(logGrid.Row, typeColNum)
        
    logGrid.ForeColorSel = ChangeLogItemColor(selectCLType)
    
End Sub