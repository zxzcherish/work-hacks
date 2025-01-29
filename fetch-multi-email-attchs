' Check the existence of the file
Function IsFileExists(ByVal strFileName As String) As Boolean
    Dim objFileSystem As Object
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    If objFileSystem.fileExists(strFileName) = True Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
End Function

Public Sub ReplaceAttachmentsToLink()
 
Dim objApp As Outlook.Application
Dim aMail As Outlook.MailItem   'Object
Dim oAttachments As Outlook.Attachments
Dim oSelection As Outlook.Selection
Dim i As Long
Dim iCount As Long
Dim sFile As String
Dim sFolderPath As String
Dim sDeletedFiles As String
 
Dim dSaved As Long
Dim dUnsaved As Long
' Get the path to your My Documents folder
sFolderPath = CreateObject("WScript.Shell").SpecialFolders(16)
On Error Resume Next
 
' Instantiate an Outlook Application object.
Set objApp = CreateObject("Outlook.Application")
 
' Get the collection of selected objects.
Set oSelection = objApp.ActiveExplorer.Selection
 
' Set the Attachment folder.
sFolderPath = sFolderPath & "\OLA转移"   # the shortform of Outlook Attachments, you can choose the filename as you like
 
dSaved = 0
dUnsaved = 0
 
' Check each selected item for attachments. If attachments exist, save them to the Temp folder and strip them from the item.
For Each aMail In oSelection
    
    ' This code only strips attachments from mail items.
    ' If aMail.class=olMail Then
    ' Get the Attachments collection of the item.
    Set oAttachments = aMail.Attachments
    iCount = oAttachments.Count
    
    If iCount > 0 Then
        
        ' Use a count down loop for removing items from a collection.
        ' Otherwise, the loop counter gets confused and only every other item is removed.
            
        For i = iCount To 1 Step -1
            
            ' Save attachment before deleting from item.
            ' Get the file name.
            sFile = oAttachments.Item(i).FileName
            If Right(sFile, 4) <> ".jpg" Or Right(sFile, 4) <> ".png" Then
                
                sMailSbj = aMail.Subject
                sFileSndr = aMail.Sender
                sFileSndr = Left(aMail.Sender, InStr(1, sFileSndr, "(") - 1)
                
                ' sFileTime = Format(aMail.CreationTime, "yyyy-mm-dd_hh-mm-ss")
                sFileDate = Format(aMail.CreationTime, "yyyy-mm-dd")

                ' this is the way I prefer my file name saved as, contains Sender, Date, Subject of the email, and the name of the file.
                ' as attachements with a same name but in diff versions can be sent to me for several times
                ' you can always change it as you prefered ;)
                sFile = sFileSndr + "_" + sFileDate + "_【"  + sMailSbj + "】_" + sFile

                ' replace the marks that are not allowed
                sFile = Replace(sFile, "：", "_")
                sFile = Replace(sFile, ":", "_")
                 
                ' combine with the path to the Temp folder.
                sFile = sFolderPath & "\" & sFile
                 
                ' save the attachment as a file.
                oAttachments.Item(i).SaveAsFile sFile
                 
                 ' Check if the attachment has been saved at the dir
                 If IsFileExists(sFile) = True Then
                     ' Delete the attachment.
                     oAttachments.Item(i).Delete
                 Else
                 ' If the file has not been saved (propably because of the file name is improper)
                     MsgBox (sMailSbj + "的附件未能被提取")
                     dUnsaved = dUnsaved + 1
                 End If
                 
                'write the save as path to a string to add to the message
                'check for html and use html tags in link
                If aMail.BodyFormat <> olFormatHTML Then
                    sDeletedFiles = sDeletedFiles & vbCrLf & "<file://" & sFile & ">"
                Else
                    sDeletedFiles = sDeletedFiles & "
" & "<a href='file://" & sFile & "'>" & sFile & "</a>"
                End If
                
            End If
     
        Next i
        'End If
                
          ' Adds the filename string to the message body and save it
          ' Check for HTML body
        If aMail.BodyFormat <> olFormatHTML Then
            aMail.Body = aMail.Body & vbCrLf & _
            "The file(s) were saved to " & sDeletedFiles
        Else
            aMail.HTMLBody = aMail.HTMLBody & "<p>" & _
            "The file(s) were saved to " & sDeletedFiles & "</p>"
        End If
          
        aMail.Save
        'sets the attachment path to nothing before it moves on to the next message.
        sDeletedFiles = ""
        
        End If
    dSaved = dSaved + 1
    Next 'end aMail
    
    MsgBox (dSaved + "封邮件的附件被尝试提取；其中，" + dUnsaved + "个附件未能被提取")
    
ExitSub:
Set oAttachments = Nothing
Set aMail = Nothing
Set oSelection = Nothing
Set objApp = Nothing
End Sub
