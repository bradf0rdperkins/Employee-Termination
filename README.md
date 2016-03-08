# Employee-Termination
Powershell script that performs multiple tasks after an employee leaves the company

This is my first Powershell script that we needed at our company to simplify the account termination process. 

It does the following:
- Sets the Recipient Limit to 0 in Exchange
- Disables OWA in Exchange
- Disables ActiveSync in Exchange
- Disables AD Account
- Changes the description to the date disabled (will be used later to move files, archive email, and delete account at specified time)
- Moves the account into the Disabled Users OU
- Removes the User from all groups and saves a log file of groups removed
- Gives the defined manager permissions to the user's mailbox
- Emails manager with details
