# famebvd
- scraps business database fame with selenium in python
- login with selenium is dependent on your login method, your institution entry page etc. for me it is logging into microsoft 365 account
- with my account i need to log in with either verification code (won't work w selenium) or authentication app. so I got authenticator app  all set up in the account in advance. when a log in attempt happens just approve from your device.
- raw export from famebvd is incredibly unfriendly for analysis so I made an upload script fameupload.py to manipulate the format and then upload to GBQ for further analysis will only work with specific set of export so meh not sure how useful it is for others

