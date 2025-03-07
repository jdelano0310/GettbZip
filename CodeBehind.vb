[Description("")]
[FormDesignerId("922804C1-F676-46FB-B8D1-EDA18513F6F0")]
[PredeclaredId]
Class Form1
    
    Dim settingsFileName As String
    
    Dim fso As Scripting.FileSystemObject
    Dim loadingSettingsFromFile As Boolean
    Dim githubReleasesPage As HTMLDocument
    Dim currentInstalledTBVersion As Integer
    Dim latestVersion As String
    
    Sub New()
    End Sub

    ' use a record to hold the change log information
    Type changeLogRecType
        tBVersion As String * 4
        clType As String * 11
        clText As String * 150
    End Type

    Private Function OptionSelection() As Integer
        
        If optOpenFolder.Value Then
            OptionSelection = 1
            
        ElseIf optOpenZip.Value Then
            OptionSelection = 2
            
        ElseIf optInstallTB.Value Then
            OptionSelection = 3
        End If
                
    End Function
    
    Private Sub GetCurrentTBVersion()
        
        ' attempt to find the version number of twinBasic in use
        Dim fileWithVersionInfo As String = txttBLocation.Text & "ide\build.js"
        Dim versionIndicator As String = "BETA"
        Dim fileContents As String
        Dim tempString As String
        
        If Not fso.FileExists(fileWithVersionInfo) Then
            lblVersion.Caption = "Version File Missing"
            Exit Sub
        End If
        
        ' open the file designated as the one with the version number
        With fso.OpenTextFile(fileWithVersionInfo, ForReading)
            fileContents = .ReadLine()
            .Close()
        End With
        
        ' parse the text for the version number
        tempString = Mid(fileContents, InStr(fileContents, versionIndicator))
        currentInstalledTBVersion = Int(Mid(tempString, Len(versionIndicator) + 1, 4))
        
        lblVersion.Caption = "Current Version: " & CStr(currentInstalledTBVersion)
        DoEvents()
        
    End Sub
    
    Private Function askForFolder(dlgCaption As String, Optional currentFolder As String) As String
        
        ' open the select folder form and return the selection (if any)
        Dim frmSelFolder As New frmSelectFolder
        
        With frmSelFolder
            If fso.FolderExists(currentFolder) Then .selectedFolder = currentFolder
            .Caption = dlgCaption
            .Show(vbModal)
        End With
        
        Return frmSelFolder.selectedFolder
        
    End Function
    
    Private Function FoldersAreValid() As Boolean
        
        ' check to see if both folders are valid
        Return fso.FolderExists(txtDownloadTo.Text) And fso.FolderExists(txttBLocation.Text)
        
    End Function
    
    Private Sub EnableDownloadZipButton()
        ' should the download zip button be enabled?
        btnDownLoadZip.Enabled = FoldersAreValid
        
        If btnDownLoadZip.Enabled Then
            ' add the final forward slash if needed
            If Right(txtDownloadTo.Text, 1) <> "\" Then txtDownloadTo.Text += "\"
            If Right(txttBLocation.Text, 1) <> "\" Then txttBLocation.Text += "\"
        End If
    End Sub
    
    Private Sub ShowStatusMessage(statMessage As String, Optional updatePreviousStatus As Boolean = False)

        ' write the message to the listbox on the form
        If updatePreviousStatus Then
            lbStatus.List(lbStatus.ListCount - 1) += statMessage
        Else
            lbStatus.AddItem(statMessage)
        End If
        
        DoEvents()
        
    End Sub
    
    Private Function GettBParentFolder() As String
        
        Dim idx As Integer
        Dim slashCount As Integer
        
        ' loop backwards until the second \ is found - which will indicate where
        ' the parent folder for twinBASIC is
        For idx = Len(txttBLocation.Text) To 1 Step -1
            If Mid(txttBLocation.Text, idx, 1) = "\" Then slashCount += 1
            If slashCount = 2 Then Exit For
        Next
        
        ' truncate the value in the textbox holding the install folder, to get the parent folder
        GettBParentFolder = Left(txttBLocation.Text, idx)
        
    End Function
    
    Private Sub InstallTwinBasic(zipLocation As String)
        
        ' go through the steps of deleting the current files and unziping the new files
        ' to the folder that has been desgniated
        
        ' delete current files & recreate the folder
        Dim SHFileOp As SHFILEOPSTRUCT
        Dim RetVal As Long
        With SHFileOp
            .wFunc = FO_DELETE
            .pFrom = txttBLocation.Text
            .fFlags = FOF_ALLOWUNDO
        End With
        RetVal = SHFileOperation(SHFileOp)
        
        ' unzip to the twinBasic folder
        With New cZipArchive
            .OpenArchive zipLocation
            .Extract txttBLocation.Text
        End With
        
        ' check to make sure the twinBASIC folder exists after attempted installation
        If fso.FolderExists(txttBLocation.Text) Then
            MsgBox("twinBasic from " & zipLocation & " has been extracted and is ready to use.", vbInformation, "Completed")
        Else
            MsgBox("There was a problem recreating " & txttBLocation.Text & ". The parent folder and the zip file will be opened so that you can finish the process.", vbCritical, "Unable to complete")
            
            ShellExecute(0, "open", zipLocation, vbNullString, vbNullString, 1) ' open the zipfile for the user
            ShellExecute(0, "open", GettBParentFolder, vbNullString, vbNullString, 1) ' open the folder where twinBASIC is supposed to be installed.
            
            MsgBox("Going forward, you can open this utility as administrator to avoid this extra step.")
            
        End If
        
    End Sub
    
    Private Sub ProcessDownloadedZip(zipLocation As String)
    
        ShowStatusMessage "Executing Post download action"
        
        ' depending on the selection, work with the zipfile downloaded
        Select Case OptionSelection
            Case 1
                ' download only - open the download folder
                ShellExecute(0, "open", txtDownloadTo.Text, vbNullString, vbNullString, 1)
            Case 2
                ' open the zip file using the default zip client
                ShellExecute(0, "open", zipLocation, vbNullString, vbNullString, 1)
            Case 3
                InstallTwinBasic(zipLocation)
                
                ' does the user want to run twinBASIC after the update
                If chkStarttwinBASIC.Value = vbChecked Then
                    ShellExecute(0, "open", txttBLocation.Text & "\twinBASIC.exe", vbNullString, vbNullString, 1)
                End If
        End Select
        
        ShowStatusMessage " - Done", True
    End Sub
    
    Private Sub CheckForSettingsFile()
        
        ' is there a settings file?
        If fso.FileExists(settingsFileName) Then
            ' there is load the dtaa from it
            ShowStatusMessage "Loading settings file"
            loadingSettingsFromFile = True ' set a flag to indicate that the app is loading the settings file
            
            Dim settingLine As String
            With fso.OpenTextFile(settingsFileName, ForReading)
                ' Download Folder
                txtDownloadTo.Text = .ReadLine()
                txtDownloadTo.Text = Trim(Split(txtDownloadTo.Text, ":")(1)) & ":" & Split(txtDownloadTo.Text, ":")(2) ' use the data to the right of each colon
                
                ' twinBasic Folder
                txttBLocation.Text = .ReadLine()
                txttBLocation.Text = Trim(Split(txttBLocation.Text, ":")(1)) & ":" & Split(txttBLocation.Text, ":")(2) ' use the data to the right of each colon
                
                ' Action
                settingLine = .ReadLine()
                settingLine = Trim(Split(settingLine, ":")(1)) ' use the data to the right of the colon
                
                ' select the option saved in the settings file
                Select Case settingLine
                    Case 1
                        optOpenFolder.Value = True
                    Case 2
                        optOpenZip.Value = True
                    Case 3
                        optInstallTB.Value = True
                End Select
                
                ' Check for new verson on load
                settingLine = .ReadLine()
                settingLine = Trim(Split(settingLine, ":")(1)) ' use the data to the right of the colon
                chkLookForUpdateOnLaunch.Value = CLng(settingLine)

                ' start twinBASIC after update
                settingLine = .ReadLine()
                settingLine = Trim(Split(settingLine, ":")(1)) ' use the data to the right of the colon
                chkStarttwinBASIC.Value = CLng(settingLine)

                .Close()
            End With
            loadingSettingsFromFile = False
            ShowStatusMessage " - Loaded", True
            
            ' are the folders from the file valid for this PC
            Dim invalidFolderMessage As String = "The folder(s)"
            If Not fso.FolderExists(txtDownloadTo.Text) Then
                invalidFolderMessage += vbCrLf & "'" & txtDownloadTo.Text & "' "
            End If
            
            If Not fso.FolderExists(txttBLocation.Text) Then
                invalidFolderMessage += vbCrLf & "'" & txttBLocation.Text & "' "
            End If
            
            If invalidFolderMessage <> "The folder(s)" Then
                ' there are issues with the samed folder(s)
                If UBound(Split(invalidFolderMessage, "'")) > 2 Then
                    ' if there are more than 2 apostrophe then there are two folders in the string
                    invalidFolderMessage += vbCrLf & "don't exist!"
                Else
                    ' one folder doesn't exist
                    invalidFolderMessage += vbCrLf & "doesn't exist!"
                End If
                
                ShowStatusMessage "Invalid settings found - process stopped"
                
                btnDownLoadZip.Enabled = False
                MsgBox(invalidFolderMessage & " on this PC", vbCritical, "Invalid Folder Settings")
                
            End If
            chkSaveSettings.Value = vbChecked
        End If
        
    End Sub
    
    Private Function GetTagText(tagName As String, tagText As String) As String

        ' retrieve all the tags that match the requested tag type and return the element
        Dim tagList As IHTMLElementCollection
        Set tagList = githubReleasesPage.getElementsByTagName(tagName)

        Dim hTag As IHTMLElement
        Dim returnText As String = ""
        
        ' searching for a first specific tag with specific text 
        For Each hTag In tagList
            If InStr(hTag.innerText, tagText) > 0 Then
                returnText = hTag.innerText
                Exit For
            End If
        Next hTag
        
        GetTagText = returnText
        
    End Function
    

    Private Function GetChangeLogUL(versionGAP As Integer) As Variant
        ' retrieve the change log section of the page
        Dim tagList As IHTMLElementCollection
        Set tagList = githubReleasesPage.getElementsByTagName("UL")

        Dim changeLogRecArray() As changeLogRecType
        
        Dim ulIndex As Integer
        Dim liIndex As Integer
        Dim liElements As IHTMLElementCollection
        Dim changeLogCount As Integer = 0  ' get the change log for each version between the installed and newest released
        Dim forVersion As String = latestVersion
        Dim arrayIndex As Integer
        Dim newArrayMax As Integer
        
        For ulIndex = 0 To tagList.length - 1
            If tagList(ulIndex).className = "" Then
                ' the change log UL tag has no class associated with it
                Set liElements = tagList(ulIndex).getElementsByTagName("LI")  ' find the LI elements in the UL
                
                If changeLogCount = 0 Then
                    ' the first change log read the array will not have any items
                    newArrayMax = liElements.length
                Else
                    newArrayMax = UBound(changeLogRecArray) + liElements.length
                End If
                
                ReDim Preserve changeLogRecArray(newArrayMax) As changeLogRecType
                
                For liIndex = 0 To liElements.length - 1
                    ' add the change log list to the dictionary to pass it back
                    changeLogRecArray(arrayIndex).tBVersion = forVersion
                    changeLogRecArray(arrayIndex).clType = Trim(Left(liElements(liIndex).innerText, InStr(liElements(liIndex).innerText, ":") - 1))
                    changeLogRecArray(arrayIndex).clText = Trim(Mid(liElements(liIndex).innerText, InStr(liElements(liIndex).innerText, ":") + 1))
                    arrayIndex += 1
                Next
                changeLogCount += 1
                
                ' once the count of captured change logs equals the number of versions between the installed and the latest - leave the loop
                If changeLogCount = versionGAP Then Exit For
                forVersion = Int(forVersion) - 1 ' as we loop more we go back to older version numbers
            End If
        Next ulIndex
        
        GetChangeLogUL = changeLogRecArray
        
    End Function
    
    Private Sub UpdateSettingsFile()
        
        ' this will be changed to use a JSON file with more information about last usage JD 1-26-25
        ' or a separate log file for each run of this utility
        If chkSaveSettings.Value = vbChecked Then
            ' the checkbox to save the form settings has been checked.

            ' write the values to the file overwriting the old if there
            With fso.CreateTextFile(settingsFileName, True)
                .WriteLine("Download Folder: " & txtDownloadTo.Text)
                .WriteLine("twinBASIC Folder: " & txttBLocation.Text)
                .WriteLine("Action: " & OptionSelection)
                .WriteLine("Check for new version on load: " & CStr(chkLookForUpdateOnLaunch.Value))
                .WriteLine("Start twinBASIC after update: " & CStr(chkStarttwinBASIC.Value))
                .Close()
            End With
                
        Else
            ' do not save, delete any file that may exist
            If fso.FileExists(settingsFileName) Then
                fso.DeleteFile(settingsFileName)
            End If
        End If
        
    End Sub
                    
    Private Sub GetLatestInfoFromReleasesPage(Optional duringFormLoad As Boolean = False)
        
        ' go to the url https://github.com/twinbasic/twinbasic/releases
        ' extract the newest version available and download it
        
        ' get the page
        ShowStatusMessage "Retrieving Releases Page"
        
        Set githubReleasesPage = New HTMLDocument
        
        If fso.FileExists(App.Path & "\GitHubReleasesPage.html") Then
            ' if this exists, debugging is happening
            ShowStatusMessage "******* using local debug html file **************"
            With fso.OpenTextFile(App.Path & "\GitHubReleasesPage.html", ForReading)
                githubReleasesPage.body.innerHTML = .ReadAll()
                .Close()
            End With
        
        Else
            Dim httpReq As New WinHttpRequest
            httpReq.Open("GET", "https://github.com/twinbasic/twinbasic/releases")
            httpReq.Send()
            httpReq.WaitForResponse()

            githubReleasesPage.body.innerHTML = httpReq.ResponseText
            
            Set httpReq = Nothing
        End If
        
        ' find the latest version number
        ShowStatusMessage "Finding latest version available"
        Dim tagText As String
        Dim versionGAP As Integer = 0
        
        tagText = GetTagText("h2", "twinBASIC BETA")
        
        latestVersion = Trim(Right(tagText, 4))
        ShowStatusMessage " : " & latestVersion, True
        
        If CInt(latestVersion) <= currentInstalledTBVersion Then
            If duringFormLoad Then
                ShowStatusMessage "Latest version already installed"
                btnDownLoadZip.Enabled = False
            Else
                MsgBox "The version in use is newer or equal to the version available on GitHub", vbInformation, "No need to update"
            End If

            ShowStatusMessage "Process stopped"
            latestVersion = ""
            Exit Sub
        Else
            ' how many versions have been released since the current version
            versionGAP = CInt(latestVersion) - currentInstalledTBVersion
        End If
        
        ' get the change log for this version
        lblNewVersion.Caption = "Upgrading to version " & latestVersion

        If versionGAP > 1 Then
            ShowStatusMessage "Extracting change log for multiple versions"
            lblChangeLogTitle.Caption = "Displaying multiple change logs"
        Else
            ShowStatusMessage "Extracting the associated change log"
        End If
         
        Dim changeLogArray() As changeLogRecType = GetChangeLogUL(versionGAP)
        Dim itemColor As Long
        Dim arrayIndex As Integer
        Dim changeLogItem As String
        
        If versionGAP = 1 Then lblChangeLogTitle.Caption = "Changelog has " & UBound(changeLogArray) & " items."
            
        lvChangeLog.ListItems.Clear()
        For arrayIndex = 0 To UBound(changeLogArray)
            changeLogItem = changeLogArray(arrayIndex).tBVersion & " - " & changeLogArray(arrayIndex).clType & ": " & changeLogArray(arrayIndex).clText
            lvChangeLog.ListItems.Add(, , changeLogItem)
                    
            Select Case Trim(changeLogArray(arrayIndex).clType)
                Case "IMPORTANT"
                    itemColor = vbBlue
                Case "KNOWN ISSUE"
                    itemColor = vbBlack
                Case "TIP"
                    itemColor = RGB(22, 83, 126) ' blueish
                Case "WARNING"
                    itemColor = RGB(153, 0, 0)   ' dark red
                Case "fixed"
                    itemColor = RGB(56, 118, 29) ' green
                Case "added"
                    itemColor = RGB(75, 0, 130)  ' indigo
                Case "updated"
                    itemColor = RGB(0, 128, 0)   ' other green 
                Case Else
                    itemColor = vbBlack
            End Select
            
            lvChangeLog.ListItems(lvChangeLog.ListItems.Count).ForeColor = itemColor
        Next
        
        Set githubReleasesPage = Nothing
        
    End Sub
                                    
    Private Sub Form_Load()
        
        Me.Caption = "twinBASIC Installer (v0.6.2)" ' doing this here as setting it in the forms properties cause the proj to not launch
        
        ' create the file system object that will be used during different code blocks
        Set fso = New FileSystemObject
        
        settingsFileName = App.Path & "\settings.txt"
        
        ' load any settings that have been saved.
        CheckForSettingsFile

        ' this contiues to check for version info if the folders have been setup
        If Len(txtDownloadTo.Text) > 0 Then
        
            ' get the current version of twinBasic
            GetCurrentTBVersion
            
            If chkLookForUpdateOnLaunch.Value Then
                GetLatestInfoFromReleasesPage True
            End If
            
            If currentInstalledTBVersion = 0 Then
                ' the current version is missing, the build.js file that contains the version information.
                Me.Show()
                btnDownLoadZip.Enabled = (MsgBox("Unable to get your current version, would you like to download the release available anyway?", vbYesNo, "ide\build.js file missing") = vbYes)
            End If
        End If
    End Sub
    
    Private Sub btnSelectDLfolder_Click()
        
        ' allow the user to select the folder to download the zip file
        Dim downloadFolder As String = askForFolder("Select download folder", txtDownloadTo.Text)
        
        If Len(downloadFolder) > 0 Then
            ' a folder was selected
            txtDownloadTo.Text = downloadFolder
        End If
    End Sub
    
    Private Sub btnSelectTBLocation_Click()
        
        ' allow the user to select the folder where twinBasic is location
        Dim twinBasicFolder As String = askForFolder("Select twinBasic folder", txttBLocation.Text)
        
        If Len(twinBasicFolder) > 0 Then
            ' a folder was selected
            txttBLocation.Text = twinBasicFolder
        End If
    End Sub
    
    Private Sub btnDownLoadZip_Click()
       
        ' if the version info from the GetReleasesPahe is already present then don't load it again.
        If chkLookForUpdateOnLaunch.Value = vbUnchecked Or lvChangeLog.ListItems.Count = 0 Then GetLatestInfoFromReleasesPage
        
        ' this will be blank if the version isn't newer than 
        ' the user already has
        If latestVersion = "" Then Exit Sub
        
        ' use the version number to download the latest release
        ' example of the dowmload url: https://github.com/twinbasic/twinbasic/releases/download/beta-x-0641/twinBASIC_IDE_BETA_641.zip
        Dim newReleaseURL As String = "https://github.com/twinbasic/twinbasic/releases/download/beta-x-" & IIf(Len(latestVersion) = 3, "0" & latestVersion, latestVersion)
        Dim justTheFileName As String = "twinBASIC_IDE_BETA_" & latestVersion & ".zip"
        Dim localZipFileName As String = txtDownloadTo.Text & justTheFileName
        Dim downloadTheZip As Boolean = True
        
        If fso.FileExists(localZipFileName) Then
            ' the zip has been downloaded already
            downloadTheZip = MsgBox("The zip file '" & localZipFileName & "' already exists. Download it again? (if no, then the current zip will be used)", vbYesNo, "Previously Downloaded") = vbYes
            If downloadTheZip Then fso.DeleteFile(localZipFileName)
        End If

        If downloadTheZip Then
            ShowStatusMessage "Downloading twinBasic " & latestVersion
            URLDownloadToFile 0, newReleaseURL & "/" & justTheFileName, localZipFileName, 0, 0
            ShowStatusMessage " - done ", True
        End If
        
        ProcessDownloadedZip localZipFileName
        
        ShowStatusMessage "process complete"
        
    End Sub
    
    Private Sub txtDownloadTo_Change()
        ' is the form reaf to download the zip file
        EnableDownloadZipButton
    End Sub
    
    Private Sub txttBLocation_Change()
        ' is the form reaf to download the zip file
        EnableDownloadZipButton
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        
        UpdateSettingsFile
        
        Set fso = Nothing
    End Sub
    
    Private Sub lvChangeLog_Click()
        
        ' crude calc just to make the currently selected item the tooltip
        ' if it is wider than the listview
        Dim pixelsPerCharacter As Integer = 78.6
        Dim lengthOfChangeLogItem As Integer = Len(lvChangeLog.SelectedItem)
        Dim requiredWidth As Integer
        
        requiredWidth = Int(lengthOfChangeLogItem * pixelsPerCharacter)
        
        ' if the change log item text needs more room than the listview gives
        ' make the selected it the tooltip of the listview, else clear the tooltip
        If requiredWidth > lvChangeLog.Width Then
            lvChangeLog.ToolTipText = lvChangeLog.SelectedItem
        Else
            lvChangeLog.ToolTipText = ""
        End If
    End Sub
    
    Private Sub optOpenFolder_Click()
        ' if just opening the folder, you can't launch the new twinBASIC
        chkStarttwinBASIC.Value = 0
        chkStarttwinBASIC.Enabled = False
        
        ' is the form reaf to download the zip file
        EnableDownloadZipButton
    End Sub
    
    Private Sub optOpenZip_Click()
        ' if just opening the zip, you can't launch the new twinBASIC
        chkStarttwinBASIC.Value = 0
        chkStarttwinBASIC.Enabled = False
        
        ' is the form ready to download the zip file
        EnableDownloadZipButton
    End Sub
    
    Private Sub optInstallTB_Click()
        ' warn the user of the process involved in installing the latest twinBASIC version
        If Not loadingSettingsFromFile Then MsgBox("Selecting this option will delete the twinBASIC folder entirely and recreate it.", vbExclamation, "Warning")
    
        ' is the form reaf to download the zip file
        EnableDownloadZipButton
        
    End Sub

    Private Sub txttBLocation_LostFocus()
            
        ' if the folder doesn't exist, create it?
        If Not fso.FolderExists(txttBLocation.Text) Then
            ' ask the user if the folder should be created (like a first time setup)
            If MsgBox("This folder doesn't exist. Should it be created?", vbYesNo, "twinBASIC Location") = vbYes Then
                On Error Resume Next
                fso.CreateFolder(txttBLocation.Text)
                If Not fso.FolderExists(txttBLocation.Text) Then
                    MsgBox("Unable to create the folder. Try another folder name.", vbCritical, "Creation Error")
                    txttBLocation.SetFocus()
                Else
                    EnableDownloadZipButton
                End If
            End If
        End If
        
    End Sub
End Class
