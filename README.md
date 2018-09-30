# MailboxCleanup
MailboxCleanup is console C# (.Net) application which helps in cleaning the JunkFolder by reading the subject line of each email. If subject line contains APPROVED or REJECTED keyword then those emails are moved to Approved and Rejected subfolder under Inbox. If email subject line is empty or subject line doesn't have APPROVED or REJECTED keyword then those emails are moved to Ignored subfolder under Inox. 

### Application Setup:

* Replace your Exchange Account Configuration in Settings.xml
* Replace Settings.xml containing your configuration in MailboxCleanup/bin/Debug folder.

### Application Information:

1. MailboxCleanup is C# console application which loads configuration settings from Settings.xml
2. MailboxCleanup application creates a connection using the configuration present in Settings.xml and counts the number of email in Junk Email Folder.
3. If JunkFolder is empty then program will not proceed and prints a message on Console stating that "Mailbox cleanup is not required because there are no junk emails in Junk Folder".
4. If JunkFolder is not empty then program proceeds and counts the number of emails present in Junk Folder and starts the cleanup process.
5. Junk Folder cleanup process is based on Subject line. Few examples of Subject lines are listed below.
#### APPROVED Subject Line Example
* SubjectLine: `[DEV TEST ONLY]Invoice Notification - Approved Invoice - [XYZ|APPROVED|000012345678:123456:1234:121212122#123456789]`
* Email subject line gets splits into 3 parts based on `|`
* If second split matches `APPROVED` keyword then those emails are moved to Approved subfolder under Inbox.

#### REJECTED Subject Line Example
* SubjectLine: `[DEV TEST ONLY]Invoice Notification - Rejected Invoice - [XYZ|REJECTED|000012345678:123456:1234:121212122#123456789]`
* Email subject line gets splits into 3 parts based on `|`
* If second split matches `REJECTED` keyword then those emails are moved to Rejected subfolder under Inbox.


#### Subject Line without APPROVED or REJECTED keyword Example
* SubjectLine: `This is a test email`
* If subject line doesn't line doesn't contain `APPROVED` or `REJECTED` keyword then those emails are moved to Ignored subfolder under Inbox.

#### Empty Subject Line Example
* SubjectLine: ` `
* If subject line is empty then those emails are moved to Ignored subfolder under Inbox.

