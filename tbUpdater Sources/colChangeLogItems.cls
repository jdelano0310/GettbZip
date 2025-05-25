VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colChangeLogItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6", "Yes"
Attribute VB_Ext_KEY = "Collection", "clsChangeLogItem"
Attribute VB_Ext_KEY = "Member0", "clsChangeLogItem"
Attribute VB_Ext_KEY = "Top_Level", "Yes"

'local variable to hold collection
Private _col As Collection
Private _githubReleasesPage As HTMLDocument
Private _installedtBVersion As Integer
Private _versionCheckResult As String
Private _tBVersionGap As Integer
Private _latestVersion As Integer

Const githubReleasesURL = "https://github.com/twinbasic/twinbasic/releases"

Public Property Let LatestVersion(lv As Integer)
    _latestVersion = lv
End Property

Public Property Get LatestVersion() As Integer

    LatestVersion = _latestVersion
End Property

Public Property Let InstalledtBVersion(installedVersion As Integer)
    _installedtBVersion = installedVersion
End Property

Public Property Get InstalledtBVersion() As Integer

    InstalledtBVersion = _installedtBVersion
End Property

Public Property Get versionCheckResult() As String

    versionCheckResult = _versionCheckResult
End Property

Public Property Get tBVersionGap() As Integer
    tBVersionGap = _tBVersionGap
End Property

Public Function Add(tBVersion As String, clType As String, clDescription As String, Optional sKey As String) As clsChangeLogItem
    
    'create a new object
    Dim newCLI As clsChangeLogItem
    Set newCLI = New clsChangeLogItem

    'set the properties passed into the method
    newCLI.tBVersion = tBVersion
    newCLI.Type = clType
    newCLI.Notes = clDescription
    
    If Len(sKey) = 0 Then
        _col.Add newCLI
    Else
        _col.Add newCLI, sKey
    End If

    'return the object created
    Set Add = newCLI
    Set newCLI = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsChangeLogItem
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = _col(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = _col.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    _col.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = _col.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set _col = New Collection
    Set _githubReleasesPage = New HTMLDocument
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set _col = Nothing
End Sub

Public Function GetNewerReleasedVersions(Optional isDuringFormLoad As Boolean = False) As Boolean
    
    ' go to the url https://github.com/twinbasic/twinbasic/releases
    ' extract the newest version available and it's release notes
    
    GetNewerReleasedVersions = False
    
    If _installedtBVersion = 0 Then
        MsgBox "Attempt to use GetNewerReleasedVersions without first setting installed version", vbCritical, "Get Released Versions"
        Exit Function
    End If
    
    If fso.FileExists(App.Path & "\GitHubReleasesPage.html") Then
        ' if this exists, debugging is happening
        _githubReleasesPage.body.innerHTML = fsoFileRead(App.Path & "\GitHubReleasesPage.html")
        
    Else
        Dim httpReq As New WinHttpRequest
        httpReq.Open("GET", githubReleasesURL)
        httpReq.Send()
        httpReq.WaitForResponse()

        _githubReleasesPage.body.innerHTML = httpReq.ResponseText
                    
        Set httpReq = Nothing
    End If

    Dim tagText As String
    Dim latesttBVersion As Integer
    
    tagText = GetTagText("h2", "twinBASIC BETA")
    latesttBVersion = Trim(Right(tagText, 4))
 
    _latestVersion = CInt(latesttBVersion)
    
    If CInt(latesttBVersion) <= _installedtBVersion Then
        If isDuringFormLoad Then
            _versionCheckResult = "Latest version already installed"

        Else
            _versionCheckResult = "The version in use is newer or equal to the version available on GitHub"
        End If

        Exit Function
    Else
        ' how many versions have been released since the current installed version
        _tBVersionGap = CInt(latesttBVersion) - _installedtBVersion
    End If
    
    On Error Resume Next
    Dim getCLs As Boolean = GetChangeLogUL(_tBVersionGap)
    If Not getCLs Then Exit Function
    On Error GoTo 0
    
    GetNewerReleasedVersions = True
    
End Function

Private Function GetTagText(tagName As String, tagText As String) As String

    ' retrieve all the tags that match the requested tag type and return the element
    Dim tagList As IHTMLElementCollection
    Set tagList = _githubReleasesPage.getElementsByTagName(tagName)

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

Private Function GetChangeLogUL(versionGAP As Integer) As Boolean
        
    ' retrieve the change log section of the page
    
    GetChangeLogUL = False
    
    Dim changeLogCount As Integer = 0  ' get the change log for each version between the installed and newest released
    Dim forVersion As String = [_latestVersion]
    
    Dim tagList As IHTMLElementCollection
    Set tagList = _githubReleasesPage.getElementsByTagName("UL")
        
    Dim ulIndex As Integer
    Dim liIndex As Integer
    Dim liElements As IHTMLElementCollection
    Dim clType As String
    Dim clDesc As String
        
    For ulIndex = 0 To tagList.length - 1
        If tagList(ulIndex).className = "" Then
            ' the change log UL tag has no class associated with it
            Set liElements = tagList(ulIndex).getElementsByTagName("LI")  ' find the LI elements in the UL
                            
            For liIndex = 0 To liElements.length - 1
                ' add the change log list to the dictionary to pass it back
                clType = Trim(Left(liElements(liIndex).innerText, InStr(liElements(liIndex).innerText, ":") - 1))
                clDesc = Trim(Mid(liElements(liIndex).innerText, InStr(liElements(liIndex).innerText, ":") + 1))
                
                Me.Add forVersion, clType, clDesc
            Next
                        
            changeLogCount += 1
                
            ' once the count of captured change logs equals the number of versions between the installed and the latest - leave the loop
            If changeLogCount = versionGAP Then Exit For
            forVersion = Int(forVersion) - 1 ' as we loop more we go back to older version numbers
        End If
    Next ulIndex
        
    GetChangeLogUL = True
        
    ' Dim relDateIndex As Integer    ******************* playing around with extracting the release date for each version
    ' Dim relDateDict As Dictionary
    
    ' ' find the release date of teh version
    ' Set tagList = _githubReleasesPage.getElementsByTagName("relative-time")
    ' For relDateIndex = 0 To tagList.length - 1
    '     If (relDateIndex Mod 2 = 0) Then
    '         ' take the even numbered elements and save to the dictionary
    '         relDateDict.Add(forVersion, tagList(relDateIndex))
    '     End If
        
    '     If (relDateIndex / 2) = versionGAP Then Exit For
    '     forVersion = Int(forVersion) - 1 ' as we loop more we go back to older version numbers
    ' Next relDateIndex
    
    ' ' reset these for the change log items
    ' forVersion = [_installedtBVersion]
    ' changeLogCount = 0
    
    
End Function