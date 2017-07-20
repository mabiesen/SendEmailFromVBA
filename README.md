# Send Email From VBA

Send email from VBA Example.  Weak encryption included.


#### How it works
Page marked "SMTPConfig" contains the variables that are used to configure an email.

Page marked "SendEmail" allows you to enter email data and your password. Your password and from email address are linked to the config tab.

To use, type your email information into the "SendEmail" page and hit send.  You may have to tweak the gmail port on the config page if you experience an issue.

#### The Internals

###### Classes
1. Encryption - Handles all encryption and decryption with a single method call
2. SMTPValues - Holds SMTP values from the configuration page
3. EmailValues - Values used in email creation including from address, to address, subject, etc.

###### Modules
1. ButtonsWParams - Called from Encrypt and Decrypt buttons which use this module as a seguay to provide variables to the sub "sendEmail".
2. EncDec - Called from "ButtonsWParams".  If 1 provided encrypts, if -1 provided decrypts
3. SendEmail - Called from send button Contains a single sub called "sendEmail".  This module provides all of the work necessary to gather information from the excel spreadsheet for email and smtp values, and sending of information.


