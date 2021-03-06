Option Explicit

Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
 
 '32-bit API declarations
Declare Function SHGetPathFromIDList Lib "shell32.dll" _
Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) _
As Long
 
Declare Function SHBrowseForFolder Lib "shell32.dll" _
Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Sub SaveAttachments()

    Dim objOL As Outlook.Application
    Dim objMsg As Outlook.MailItem 'Object
    Dim objAttachments As Outlook.Attachments
    Dim objSelection As Outlook.Selection
    Dim i As Long
    Dim lngCount As Long
    Dim strFile As String
    Dim strFolderDir As String
    Dim strFolderpath As String
    Dim strDeletedFiles As String
    Dim Response As Variant
    
    Dim ContainsAttachements As Boolean: ContainsAttachements = False

    ' Instantiate an Outlook Application object.
    Set objOL = CreateObject("Outlook.Application")

    ' Get the collection of selected objects.
    Set objSelection = objOL.ActiveExplorer.Selection

    For Each objMsg In objSelection
        If objMsg.Attachments.Count > 0 Then
            ContainsAttachements = True
            Exit For
        End If
    Next
    
    If ContainsAttachements = False Then
        MsgBox "No Attachements Selected!"
        Exit Sub
    End If

    On Error GoTo Error1
    ' Get the path to your My Documents folder
    
    strFolderDir = GetDirectory() & "\"
        
    If strFolderDir = "\" Then
        strFolderpath = InputBox("No file selected." & Chr(10) & "Please enter the Filepath..")
        If strFolderpath = "" Then
            GoTo Response1
        Else: GoTo Response2
        End If
    End If
    
    'Prompt user for new folder
    Response = MsgBox("Would you like to create a new folder within this directory?", vbYesNo)
    
Response1:
    'if yes then contacate strFolderpath with new Folderpath.
    If Response = vbYes Then
        strFolderpath = InputBox("Please provide a new folder name..." & Chr(10) & "Type the '\' key for multiple folders")
    End If
        
    'If '\', '/' characters were found
    Do While InStr(strFolderpath, "\") > 0
                    
        strFolderDir = strFolderDir & Left(strFolderpath, InStr(strFolderpath, "\"))
        
        'Creates new Folderpath, uses conditional to prevent crashes
        If Dir(strFolderDir, vbDirectory) = "" Then MkDir strFolderDir
        
        strFolderpath = Right(strFolderpath, Len(strFolderpath) - InStr(strFolderpath, "\"))
                
    Loop

Response2:

    'Combines strFolderpath with strFolderpath with validation to prevent crashing.
    If strFolderpath <> "\" Or strFolderpath <> "" Or Right(strFolderDir, 1) <> "\" Then _
        strFolderDir = strFolderDir & strFolderpath & "\"
        
    'Creates new Folderpath, uses conditional to prevent crashes
    If Dir(strFolderDir, vbDirectory) = "" Then MkDir strFolderDir
            
    On Error Resume Next

    ' Check each selected item for attachments.
    For Each objMsg In objSelection

        Set objAttachments = objMsg.Attachments
        lngCount = objAttachments.Count

        If lngCount > 0 Then
    
        ' Use a count down loop for removing items
        ' from a collection. Otherwise, the loop counter gets
        ' confused and only every other item is removed.

            For i = lngCount To 1 Step -1
        
                ' Get the file name.
                strFile = objAttachments.Item(i).FileName
            
                ' Combine with the path to the Temp folder.
                strFile = strFolderDir & strFile
            
                ' Save the attachment as a file.
                objAttachments.Item(i).SaveAsFile strFile

            Next i
        End If
    Next
        
    Call Shell("explorer.exe" & " " & strFolderDir, vbNormalFocus)

ExitSub:

    Set objAttachments = Nothing
    Set objMsg = Nothing
    Set objSelection = Nothing
    Set objOL = Nothing
    Set Response = Nothing
Exit Sub

Error1:
    Response = MsgBox("Invalid Folder Name. Retry?", vbYesNo)
    If Response = vbYes Then
        GoTo Response1
    Else: MsgBox "Operation is aborted...", vbCritical, "Aborted"
    End If
Exit Sub

End Sub
                  
 
Function GetDirectory(Optional Msg) As String
    Dim bInfo As BROWSEINFO
    Dim path As String
    Dim r As Long, x As Long, pos As Integer
     
     '   Root folder = Desktop
    bInfo.pidlRoot = 0&
     
     '   Title in the dialog
    If IsMissing(Msg) Then
        bInfo.lpszTitle = "Select a folder."
    Else
        bInfo.lpszTitle = Msg
    End If
     
     '   Type of directory to return
    bInfo.ulFlags = &H1
     
     '   Display the dialog
    x = SHBrowseForFolder(bInfo)
     
     '   Parse the result
    path = Space$(512)
    r = SHGetPathFromIDList(ByVal x, ByVal path)
    If r Then
        pos = InStr(path, Chr$(0))
        GetDirectory = Left(path, pos - 1)
    Else
        GetDirectory = ""
    End If
End Function
