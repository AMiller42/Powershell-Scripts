# O365 Email Logs

This script is used to pull email logs from O365. It exports the logs to a CSV with the following fields:

* The date/time the email was received
* The email address it was sent from
* The IP address it was sent from
* The email address it was sent to
* The IP address it was sent to
* The subject of the email
* The status of the email
* The size of the email

The script prompts for the number of days to search through, as well as the size and location of the logs. By default, it searches the last 7 days, and saves the logs with 5000 entries each to the C:\\Users\User\Documents folder.

When you run the script and go through the prompts, a box will pop up asking you to sign into Office 365. Sign in with an admin account with sufficient permissions, and the script will start pulling the logs. Once it is finished, it will disconnect the O365 session.

# Prerequisites

This script connects to O365 with the EXO V2 module, which can be downloaded from the Powershell Gallery. Instructions for installing this module can be found [here](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module).
