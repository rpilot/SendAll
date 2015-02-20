Public Sub SendDrafts() 

Dim lDraftItem As Long
Dim myOutlook As Outlook.Application
Dim myNamespace As Outlook.NameSpace
Dim myFolders As Outlook.Folders
Dim myDraftsFolder As Outlook.MAPIFolder
'Send all items in the "Drafts" folder that have a "To" address filled in. 

'Setup Outlook 

Set myOutlook = Outlook.Application
Set myNamespace = myOutlook.GetNamespace("MAPI")
Set myFolders = myNamespace.Folders
'Set Draft Folder. This will need modification based on where it's being run. 
Set myDraftsFolder = myNamespace.GetDefaultFolder(olFolderDrafts) 
'Loop through all Draft Items 

For lDraftItem = myDraftsFolder.Items.Count To 1 Step -1 

'Check for "To" address and only send if "To" is filled in. 

If Len(Trim(myDraftsFolder.Items.item(lDraftItem).To)) > 0 Then 

'Send Item 

myDraftsFolder.Items.item(lDraftItem).Send 

End If
Next lDraftItem 

'Clean-up 

Set myDraftsFolder = Nothing
Set myNamespace = Nothing
Set myOutlook = Nothing 

End Sub
