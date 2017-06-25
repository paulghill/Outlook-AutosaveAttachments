'Copyright 2017 Paul Hill
'
'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
'
'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.

Option Explicit

'Save location for attachments.  Network locations will work also.
'Modify AttachmentSaveLocation to a location the system can read and write to.
Const AttachmentSaveLocation = "C:\"
'Save location for the folder id file.  Modify to a location the system can read and write to.
Const FolderIdSaveLocation = "C:\FolderId.txt"
'The folder used to scan for new messages to fire event to check if message attachments need to be saved.
'Run the GetFolderID subroutine to get entry ID of a folder.
Const FolderId = ""

'Sets the event to watch a folder for incoming items
Public WithEvents SaveAttachmentsFolder As Outlook.Items

'Sets event to watch a folder for new items
Private Sub Application_Startup()
    'Comment the line below out if you wish to use a specific
    Set SaveAttachmentsFolder = Application.Session.GetDefaultFolder(olFolderInbox).Items
    
    'If you uncomment this line, make sure you have the constant FolderId set to the
    'correct value and comment the Set SaveAttachmentsFolder line above.  Only one of 
    'these two code lines should be used.
    'Set SaveAttachmentsFolder = Application.Session.GetFolderFromID(FolderId).Items
End Sub

Private Function GetReferenceId() As String
    Dim referenceId As String
    Dim script As Object
    
    'Set object to library we need to use
    Set script = CreateObject("Scriptlet.TypeLib")
    
    'Create the GUID and only return the 36 characters
    referenceId = Mid$(script.GUID, 2, 36)
    
    'Uncomment to remove hypens.  This can be commented out if you wish to retain the hypens.
    'referenceId = Replace(referenceId, "-", vbNullString, Compare:=vbTextCompare)
    
    GetReferenceId = referenceId
End Function

'Customize this subroutine as needed for your use.
'If this returns true, then current new mail has attachment that will need to be saved
Private Function IsAttachmentEmail(Item As Object) As Boolean
    Dim output As Boolean
    
    'Ensure the item being passed is a MailItem
    If TypeOf Item Is Outlook.MailItem Then
        'Ensure the MailItem has attachments
        If Item.attachments.count > 0 Then
            'Place custom checks here, the previous checks should remain in place.
            If InStr(Item.subject, "") > 0 Then
                output = True
            Else
                output = False
            End If
            'End custom checks
        Else
            output = False
        End If
    Else
        output = False
    End If
    
    IsAttachmentEmail = output
    
End Function

Private Sub SaveAttachments(ByVal Item As Object)
    
    'If the IsAttachmentEmail is true and has attachments
    If IsAttachmentEmail(Item) Then
        Dim fileName As String
        Dim attachment As Outlook.attachment
        
        'Loop through all attachments
        For Each attachment In Item.attachments
            'Create the save location and file name for each attachment
            fileName = AttachmentSaveLocation & attachment.filename
            
            'Save the attachment
            attachment.SaveAsFile fileName
        Next
    End If
    
End Sub
 
'This subroutine is used every time the folder adds an item, i.e. a new message
Private Sub SaveAttachmentsFolder_ItemAdd(ByVal Item As Object)
    SaveAttachments Item
End Sub

'Use this subroutine to get the folder id of an Outlook folder.
Public Sub GetFolderId()
    Dim referenceId As String
    
    'This will bring up a folder picker for Outlook.  Select the folder you want the EntryId for
    referenceId = Application.Session.PickFolder.EntryID
    
    'The EntryId is a very long string.  This will save the id to a file to copy and paste back into
    'the script to use a different folder for scanning and saving attachments
    Open FolderIdSaveLocation For Output As 1
        Print #1, referenceId
    Close #1
End Sub

'Allows manual saving of attachments from selected MailItems that
'meet the criteria defined in the SaveAttachments subroutine
Public Sub SaveAttachmentsFromSelection()
    Dim o As Object
    
    'Cycle through each selection
    For Each o In Application.ActiveExplorer.selection
        'Send the object to SaveAttachments
        SaveAttachments o
    Next o
End Sub