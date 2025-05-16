Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

'Categorize Sent Items

'store it in different folder

'Place in ThisOutlookSession

 

    If TypeOf Item Is Outlook.MailItem And Len(Item.Categories) = 0 Then

        Set Item = Application.ActiveInspector.CurrentItem

        Item.ShowCategoriesDialog

        Item.UnRead = False

       

        

        'Setting the SaveSentMessageFolder property for the new message to a custom folder

        'This will let me organize things by category, in one folder.

            Set InboxFolder = Application.Session.GetDefaultFolder(olFolderInbox)

            Set desFolder = InboxFolder.Folders("_processed email")

            Set Item.SaveSentMessageFolder = desFolder

         

          '# Not using this prompt for now, keeping code for reference.

          'Prompt for destination folder, and then save the message to that folder.

            'Set desFolder = Application.Session.PickFolder

            'Set Item.SaveSentMessageFolder = desFolder

           

            'Debugging

            'MsgBox (desfolder)

         

    End If

End Sub

 