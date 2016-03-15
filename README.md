# PowerShellSetAutomaticReplies
A powershell script to set automatic replies on an exchange mailbox.

### Features:
* Easy to use GUI.
* Authentication required.
* Set multiple automatic replies without restarting the script.

### Configuration
The following lines need to be edited to your liking and environment.
* Line 19 allows you to change the directory that is temp made to format the user list. This needs to be writable.
* Line 20 allows you to change the location of the user list whilst formatting users. This needs to be writable.
* Line 124 allows you to change the directory that is temp made to format the user list. This needs to be writable.
* Line 125 allows you to change the location of the user list whilst formatting users. This needs to be writable.
* Line 145 allows you to change the DC and OU to search for users - this will need changing to your environment.

### Prerequisites
* RSAT (Remote Server Administration Tools) are needed for this script, they can be downloaded from Microsoft.
* Execution Policy needs to be set to Unrestricted to do this open PowerShell and type "Set-ExecutionPolicy – Unrestricted".

Feel free to submit fork and pull requests.
http://gravzy.com/
