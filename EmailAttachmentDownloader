Private WithEvents Items As Outlook.Items
Private Sub Application_Startup()
    Dim olNs As Outlook.NameSpace
    Dim Inbox As Outlook.MAPIFolder
    Dim Filter As String
    
    'Makes a DASL query based on subject and if the email has an attachment
    Filter = "@SQL=" & Chr(34) & "urn:schemas:httpmail:from" & _
                       Chr(34) & " Like '%@rusalamerica.com%' OR " & _
                       Chr(34) & "urn:schemas:httpmail:from" & _
                       Chr(34) & " Like '%@sellerslogistics.com%' OR " & _
                       Chr(34) & "urn:schemas:httpmail:from" & _
                       Chr(34) & " Like '%@eramet.com%' OR " & _
                       Chr(34) & "urn:schemas:httpmail:from" & _
                       Chr(34) & " Like '%@millerandco.com%' OR " & _
                       Chr(34) & "urn:schemas:httpmail:from" & _
                       Chr(34) & " Like '%@boscus.com%' AND " & _
                       Chr(34) & "urn:schemas:httpmail:hasattachment" & _
                       Chr(34) & "=1"
    
    Set olNs = Application.GetNamespace("MAPI")
    Set Inbox = olNs.GetDefaultFolder(olFolderInbox)
    Set Items = Inbox.Items.Restrict(Filter)
End Sub

Private Sub Items_ItemAdd(ByVal Item As Object)
    If TypeOf Item Is Outlook.MailItem Then
        Dim Atmtname As String
        Dim FilePath As String
            FilePath = "C:\Users\kyle.conrad\Documents\Kyle Conrad\VBA Scripts\Downloads\"
            
        Dim Atmt As Attachment
        For Each Atmt In Item.Attachments
            Atmtname = FilePath & Atmt.FileName
            Debug.Print Atmtname
            Atmt.SaveAsFile Atmtname
        Next Atmt
End Sub
