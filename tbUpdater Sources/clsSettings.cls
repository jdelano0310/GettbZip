VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
    Persistable = 0  'NotPersistable
    DataBindingBehavior = 0  'vbNone
    DataSourceBehavior  = 0  'vbNone
    MTSTransactionMode  = 0	'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' add your member fields and procedures here
Private _downloadFolder As String
Private _twinBASICFolder As String
Private _postDownloadAction As Integer
Private _checkForNewVersionOnLoad As Integer
Private _starttwinBASICAfterUpdate As Integer
Private _logActivity As Integer
Private _settingsFileLoaded As Boolean
Private _settingsChanged As Boolean
Private _saveSettingsOnExit As Integer

Private Sub Class_Initialize()
    'creates the collection when this class is created
    _settingsFileLoaded = LoadSettings()
End Sub

Public Function LoadSettings() As Boolean
    
    Dim settingsContent As String = fsoFileRead(App.Path & "\settings.txt")
    If settingsContent = "Failed fsoFileRead" Then
        Return False
    End If
    
    Dim settingsLines() As String
    Dim settingsLine As String
    Dim settingName As String
    Dim settingValue As String
    
    settingsLines = Split(settingsContent, vbCrLf) ' create an array of the settings values found in the file
    
    ' loop the array to fill in the form's controls 
    For Each settingsLine In settingsLines
        If Len(Trim(settingsLine)) > 0 Then
            
            settingName = Split(settingsLine, ": ")(0)
            settingValue = Split(settingsLine, ": ")(1)
            
            Select Case settingName
                Case "Download Folder"
                    _downloadFolder = settingValue
                Case "twinBASIC Folder"
                    _twinBASICFolder = settingValue
                Case "Action"
                    _postDownloadAction = CInt(settingValue)
                Case "Save Settings On Exit"
                    _saveSettingsOnExit = CInt(settingValue)
                Case "Check for new version on load"
                    _checkForNewVersionOnLoad = CInt(settingValue)
                Case "Start twinBASIC after update"
                    _starttwinBASICAfterUpdate = CInt(settingValue)
                Case "Log updates"
                    _logActivity = CInt(settingValue)
            End Select
        End If
    Next
    
    _settingsChanged = False
    Return True
    
End Function

Public Function SaveSettings() As Boolean
    
    ' write the values to the file overwriting the old if there
    With fso.CreateTextFile(App.Path & "\settings.txt", True)
        .WriteLine("Download Folder: " & _downloadFolder)
        .WriteLine("twinBASIC Folder: " & _twinBASICFolder)
        .WriteLine("Action: " & _postDownloadAction)
        .WriteLine("Save Settings On Exit: " & CStr(_saveSettingsOnExit))
        .WriteLine("Check for new version on load: " & CStr(_checkForNewVersionOnLoad))
        .WriteLine("Start twinBASIC after update: " & CStr(_starttwinBASICAfterUpdate))
        .WriteLine("Log updates: " & CStr(_logActivity))
        .Close()
    End With
    
End Function

Public Property Get SettingsFileLoaded() As Boolean
    SettingsFileLoaded = _settingsFileLoaded
End Property

Public Property Let DownloadFolder(ByVal vData As String)
    _downloadFolder = vData
    _settingsChanged = True
End Property

Public Property Get DownloadFolder() As String
    DownloadFolder = _downloadFolder
End Property

Public Property Let twinBASICFolder(ByVal vData As String)
    _twinBASICFolder = vData
    _settingsChanged = True
End Property

Public Property Get twinBASICFolder() As String
    twinBASICFolder = _twinBASICFolder
End Property

Public Property Let PostDownloadAction(ByVal vData As Integer)
    _postDownloadAction = vData
    _settingsChanged = True
End Property

Public Property Get PostDownloadAction() As Integer
    PostDownloadAction = _postDownloadAction
End Property

Public Property Let CheckForNewVersionOnLoad(ByVal vData As Integer)
    _checkForNewVersionOnLoad = vData
    _settingsChanged = True
End Property

Public Property Get CheckForNewVersionOnLoad() As Integer
    CheckForNewVersionOnLoad = _checkForNewVersionOnLoad
End Property

Public Property Let StarttwinBASICAfterUpdate(ByVal vData As Integer)
    _starttwinBASICAfterUpdate = vData
    _settingsChanged = True
End Property

Public Property Get StarttwinBASICAfterUpdate() As Integer
    StarttwinBASICAfterUpdate = _starttwinBASICAfterUpdate
End Property

Public Property Let LogActivity(ByVal vData As Integer)
    _logActivity = vData
    _settingsChanged = True
End Property

Public Property Get LogActivity() As Integer
    LogActivity = _logActivity
End Property

Public Property Get SettingsChanged() As Boolean
    SettingsChanged = _settingsChanged
End Property

Public Property Let SaveSettingsOnExit(ByVal vData As Integer)
    _saveSettingsOnExit = vData
End Property

Public Property Get SaveSettingsOnExit() As Integer
    SaveSettingsOnExit = _saveSettingsOnExit
End Property
