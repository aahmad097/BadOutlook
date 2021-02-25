# Bad Outlook

A simple PoC which leverages the Outlook Application Interface (COM Interface) to execute shellcode on a system based on a specific trigger subject line. 

By utilizing the `Microsoft.Office.Interop.Outlook` namespace, developers can represent the entire Outlook Application (or at least according to [Microsoft](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.application?view=outlook-pia)). This means that the new application should  be able to do anything from reading emails (yes this also includes archives, trash, etc.) to sending them out.

Building on the millions of pre-existing C# shellcode loaders, an email with a trigger subject line and base64 encoded shellcode in the body can be sent to the host with a weaponized instance of this program. The program will then read the  email and execute the shellcode embedded in the email. 

## Additional Notes:
- This can be used to build an Entire C2 Framework that relies on E-Mails as a mean of communication (Where the Implant never speaks to the internet directly)
- There does appear to be a security warning which informs the user of an application attempting to access Outlook data
    - This can be turned off with when an administrator modifies via [registry.](https://docs.microsoft.com/en-us/outlook/troubleshoot/security/a-program-is-trying-to-send-an-email-message-on-your-behalf)
    - Minor testing showed that Injecting this process into an Outlook client does not cause the alert to appear (Additional testing would be much appriciated <3)

## PoC

Applicaiton Polling Outlook for Trigger

![system schema](https://github.com/S4R1N/BadOutlook/blob/master/PoC/Checks.png)

Trigger Email With Shellcode Creation

![system schema](https://github.com/S4R1N/BadOutlook/blob/master/PoC/EmailGeneration.png)

Email Recived By Outlook Client

![system schema](https://github.com/S4R1N/BadOutlook/blob/master/PoC/EmailReceived.png)

Shellcode Execution by BadOutlook Application

![system schema](https://github.com/S4R1N/BadOutlook/blob/master/PoC/Trigger.png)
