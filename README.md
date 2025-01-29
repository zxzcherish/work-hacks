# Working hacks useful for me

## make-attachments-organised-again
Hacks for daily work: arrangement for email attachments, move attachments (in physical) from a batch of emails to a file path you like.

fetch-multi-email-attchs writes in VBA. 

The origin of the code is that a friend of mine looked forward to removing all the attachments from the emails, but keeping the original email, so that communication history could be traced back, while the storage that Outlook email takes was less.

Afterwards, I felt like moving these attachements out could be a better idea, so that we will be able to compare those version, keeping the "final_v12_revised" version with the lastest date and deleting all the others in case those are delivery or important docs.

Target user:
Outlook email user who has authorisation to Microsoft Macro.

Target of the code:
1. Fetch attachements out from the chosen emails, save them to a path you determined
2. Remove the attachments from emails, and leave the link of the file by path and filename in the original email

Using method:
1. Create a module
2. Paste the code into it, change the directory name if you like
3. Multi-choose your emails
4. Run the VBA module
