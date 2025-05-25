VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colHistoryLogItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6", "Yes"
Attribute VB_Ext_KEY = "Collection", "clsHistoryLogItem"
Attribute VB_Ext_KEY = "Member0", "clsHistoryLogItem"
Attribute VB_Ext_KEY = "Top_Level", "Yes"

'local variable to hold collection
Private _logLines() As String
Private _logLineCLI As clsChangeLogItem
Private _historyLogDatesArray() As String
Private _historyLogCol As Collection
Private _logFileLoaded As Boolean

Public Property Get HistoryLogDates() As Variant
    HistoryLogDates = _historyLogDatesArray()
End Property

Public Function Add(hli As clsHistoryLogItem, Optional sKey As String) As clsHistoryLogItem
    
    'create a new object
    Dim newHLI As New clsHistoryLogItem
        
    'set the properties passed into the method
    If Len(sKey) = 0 Then
        _historyLogCol.Add hli
    Else
        _historyLogCol.Add hli, sKey
    End If

    'return the object created
    Set Add = hli
    Set hli = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsHistoryLogItem
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = _historyLogCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = _historyLogCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    _historyLogCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = _historyLogCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set _historyLogCol = New Collection
    Set _logLineCLI = New clsChangeLogItem

End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated

    Set _historyLogCol = Nothing
    Set _logLineCLI = Nothing
    
End Sub

Public Function LoadLog() As Boolean
    
    Dim LogFileContents As String = fsoFileRead(App.Path & "\log.txt")
    If LogFileContents = "Failed fsoFileRead" Then
        Return False
    End If
    
    Dim LogLines() As String
    Dim LogLine As String
    Dim LogDateTime As String
    Dim settingValue As String
    Dim lastDateRead As String = ""
    Dim tbVersion As String
    Dim logMessage As String
    Dim historyLogItem As clsHistoryLogItem
    Dim uniqueLogDateCount As Integer = -1
        
    LogLines = Split(LogFileContents, vbCrLf)
    
    For Each LogLine In LogLines
        
        If Len(LogLine) = 0 Then Exit For
        
        Set historyLogItem = New clsHistoryLogItem
        historyLogItem.LogCLI = New clsChangeLogItem
        
        ' check for a version number  - if found then the line is the changelog data
        ' of the applied tB version
        LogDateTime = Left(LogLine, 20)
        
        tbVersion = Mid(LogLine, 23, 4)
        
        If lastDateRead <> Format(CDate(LogDateTime), "MM/dd/yy") Then
            If Len(lastDateRead) > 0 Then
                ' if there is a lastDateRead, then add the date to the date collection
                ' used to fill teh combobox on the Log tab
                uniqueLogDateCount = uniqueLogDateCount + 1
                ReDim Preserve _historyLogDatesArray(uniqueLogDateCount)
                _historyLogDatesArray(uniqueLogDateCount) = lastDateRead
            End If
            lastDateRead = Format(CDate(LogDateTime), "MM/dd/yy")
        End If
        
        ' if this character position is empy then skip the line
        If Mid(LogLine, 23, 1) <> " " Then
            If InStr(tbVersion, "/") > 1 Then Stop
            If Val(tbVersion) = 0 Then
                ' no version number, this is just a log message
                logMessage = Mid(LogLine, 23)
                                
                historyLogItem.LogDateTime = LogDateTime
                historyLogItem.LogMessage = logMessage
                historyLogItem.LogCLI.tBVersion = 0
                historyLogItem.LogCLI.Type = ""
                historyLogItem.LogCLI.Notes = ""
                                
            Else
                
                historyLogItem.LogCLI.tBVersion = CInt(tbVersion)
                historyLogItem.LogCLI.Type = Mid(LogLine, 30, 11)
                historyLogItem.LogCLI.Notes = Mid(LogLine, 43)
                historyLogItem.LogMessage = ""
                historyLogItem.LogDateTime = LogDateTime
                
            End If
            
            Me.Add historyLogItem
        End If
                
        Set historyLogItem = Nothing
    Next LogLine
    
    ' include the last date encountered
    uniqueLogDateCount = uniqueLogDateCount + 1
    ReDim Preserve _historyLogDatesArray(uniqueLogDateCount)
    _historyLogDatesArray(uniqueLogDateCount) = lastDateRead
    
    Return True
    
End Function

