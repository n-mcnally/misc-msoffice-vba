Attribute VB_Name = "SaveAllAttachments" ' remove if not importing 

'=====================================================================
'Description: Outlook macro to save all pictures that are shown in
'         the selected message bodies (embedded) as well as all
'         other attachments in their original file format.
'
'         If multiple messages are selected the file names will be
'         prefixed using the source messages subject line.
'
'         A modified version of Robert Sparnaaij's script with
'         support for multiple message processing.
'
'         See souce link below for installation instructions.
'
'Important!   This macro requires:
'             - a reference to Microsoft Shell and Automation
'               - In VBA Editor: Tools-> References...
'             - macros enabled
'               - In Trust Center: Options -> Trust Center ...
'       
'Version : 1.1
'Authors : Robert Sparnaaij & Nathan Mcnally
'Source  : https://www.howto-outlook.com/howto/saveembeddedpictures.htm
'
'=====================================================================

Private Const INITIAL_FOLDER_PATH As String = ""

Private Const BIF_RETURNONLYFSDIRS As Long = &H1

'Checks user selected path for unsupported targets
Function BrowseFolder(Optional Caption As String, _
    Optional InitialFolder As String) As String

    Dim SH As Shell32.Shell
    Dim F As Shell32.Folder
    Dim WshShell As Object
    
    'Default to user directory if initial value not set
    If InitialFolder = "" Then
        InitialFolder = Environ("USERPROFILE") & "\"
    End If
        

    Set SH = New Shell32.Shell
    Set F = SH.BrowseForFolder(0&, Caption, BIF_RETURNONLYFSDIRS, InitialFolder)
    Set WshShell = CreateObject("WScript.Shell")

    If Not F Is Nothing Then
        'Check for special folders that don't always return their full path
        Select Case F.Title
            Case "Desktop"
                BrowseFolder = WshShell.SpecialFolders("Desktop")
            Case "My Documents"
                BrowseFolder = WshShell.SpecialFolders("MyDocuments")
            Case "My Computer"
                MsgBox "Invalid selection", vbCritical + vbOKOnly, "Error"
                Exit Function
            Case "My Network Places"
                MsgBox "Invalid selection", vbCritical + vbOKOnly, "Error"
                Exit Function
            Case Else
                BrowseFolder = F.Items.Item.Path
        End Select
   End If
   
   'Cleanup
   Set SH = Nothing
   Set F = Nothing
   Set WshShell = Nothing

End Function

'Remove illegal characters from string
Function ReplaceCharsForFileName(sSubject As String, sChr As String) As String
    Dim output As String

    output = sSubject
    output = Replace(output, "'", sChr)
    output = Replace(output, "'", sChr)
    output = Replace(output, "*", sChr)
    output = Replace(output, "/", sChr)
    output = Replace(output, "\", sChr)
    output = Replace(output, ":", sChr)
    output = Replace(output, "?", sChr)
    output = Replace(output, Chr(34), sChr)
    output = Replace(output, "<", sChr)
    output = Replace(output, ">", sChr)
    output = Replace(output, "|", sChr)
    output = Replace(output, " ", "_")
    
    ReplaceCharsForFileName = output
End Function

'Get current message item attachments and save to disk
Function SaveMessageAttachmentsToDisk(MessageItem As Object, FolderPath As String, _
    ItemCount As Integer) As Integer

    SaveMessageAttachments = 0
    
    'Retrieve all attachments from the selected item
    Dim colAttachments As Outlook.Attachments
    Dim objAttachment As Outlook.Attachment
    Set colAttachments = MessageItem.Attachments
    
    Dim DateStamp As String
    Dim MyFile As String
    
    'Starts file name with subject line if multple items selected
    Dim SubjectLine As String
    
    If ItemCount > 1 Then
        SubjectLine = Left(MessageItem.Subject, 25)
        SubjectLine = "(" & ReplaceCharsForFileName(SubjectLine, "-") & ") "
    Else
        SubjectLine = ""
    End If

    For Each objAttachment In colAttachments
        MyFile = objAttachment.FileName
        DateStamp = Format(MessageItem.CreationTime, "-yymmdd_hhnn")
        intPos = InStrRev(MyFile, ".")
        
        'Add timestamp before file extention
        If intPos > 0 Then
            MyFile = SubjectLine & Left(MyFile, intPos - 1) & DateStamp & Mid(MyFile, intPos)
        Else
            MyFile = SubjectLine & MyFile & DateStamp
        End If
        
        objAttachment.SaveAsFile (FolderPath & "\" & MyFile)
    Next
    
    'Cleanup
    Set objAttachment = Nothing
    Set colAttachments = Nothing
    
    SaveMessageAttachments = 1
End Function

'Diplsay success alert and optionally open target in File Explorer
Function SuccessfullNotication(FolderPath As String)
    Dim msgBoxResponse As VbMsgBoxResult

    msgBoxResponse = MsgBox("Successfully saved attachment(s) to disk." & vbNewLine & vbNewLine & "Open in File Explorer?", vbInformation + vbYesNo, "Save all Attachments")

    'Don't open in file explorer
    If msgBoxResponse = vbYes Then
        Shell "C:\WINDOWS\explorer.exe """ & FolderPath & "", vbNormalFocus
    End If

End Function

'Macro entry point
Sub SaveAllAttachments()
    'Get all selected items
    Dim objOL As Outlook.Application
    Dim objSelection As Outlook.Selection
    Dim objItem As Object
    Set objOL = Outlook.Application
    Set objSelection = objOL.ActiveExplorer.Selection

    'Make sure at least one item is selected
    If objSelection.Count < 1 Then
       Response = MsgBox("Please select at least one item", vbExclamation, "Save all Attachments")
       Exit Sub
    End If

    'Prompt user for output directory selection
    Dim FolderPath As String
    FolderPath = BrowseFolder("Select a folder", INITIAL_FOLDER_PATH)

    If FolderPath = "" Then
        Response = MsgBox("Please select a folder. No items were saved", vbExclamation, "Save all Attachments")
       Exit Sub
    End If

    Dim CurrentItemIndex As Integer
    Dim CurrentItem As Object

    For CurrentItemIndex = 1 To objSelection.Count Step 1
        Set CurrentItem = objSelection.Item(CurrentItemIndex)
        Response = SaveMessageAttachmentsToDisk(CurrentItem, FolderPath, objSelection.Count)
    Next CurrentItemIndex

    'Success and open in File Explorer dialog
    Response = SuccessfullNotication(FolderPath)
    
    'Cleanup
    Set CurrentItem = Nothing
    Set objOL = Nothing
    Set objSelection = Nothing
    Set objItem = Nothing
End Sub