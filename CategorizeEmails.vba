'=================
' DISCLAIMER:  
'  This is some code that I wrote that I find helpful for my purposes. 
'  It may not work for you and like any code posted by any 3rd party, there is always 
'  the potential for unintended harm. 
'  Use at your own risk. 
'  No warranty or suitability for purpose is implied or intended.  
' 
'
' This macro will prompt you to categorize the current outgoing email
' and place it a folder called '_ProcessedMail' 
' Author: Sredoje Vakareskov
' 
' note: depending on your org restrictions, you may have to sign the mmacro.  
' see other docs in this repo for some help with that. 
' 
' ========================================



Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

'Categorize Sent Items
'store it in different folder
'Place in ThisOutlookSession
    If TypeOf Item Is Outlook.MailItem And Len(Item.Categories) = 0 Then

        Item.ShowCategoriesDialog
        Item.UnRead = False
        
        'Setting the SaveSentMessageFolder property for the new message to a custom folder

        'This will let me organize things by category, in one folder.
            Set InboxFolder = Application.Session.GetDefaultFolder(olFolderInbox)
            Set desfolder = InboxFolder.Folders("_ProcessedMail")
            Set Item.SaveSentMessageFolder = desfolder
            
            'MsgBox (Item.SaveSentMessageFolder)
          
          '# Not using this prompt for now, keeping code for reference.
          'Prompt for destination folder, and then save the message to that folder.
            'Set desFolder = Application.Session.PickFolder
            'Set Item.SaveSentMessageFolder = desFolder
            'Debugging
            'MsgBox (desfolder)

    End If

End Sub
