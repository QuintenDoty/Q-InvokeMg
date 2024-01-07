# Description

The reason that this repository was created is I found while using ````Get-MgUser```` & ````Update-MgUser ````, I ran into certain data points that I could not gather or update. To mitigate this, I began using Invoke-MgGraphRequest. I wanted to be able to share these methods with colleagues and here we are. Soon enough, it more than anything became a learning experience on how to create PowerShell functions and modules and get it published. Crazy how easy it is to just get a package up on PSGallery.

Here is PSGallery Link: https://www.powershellgallery.com/packages/Q-InvokeMg

# How to use

This can be downloaded directly from powershell using the following command
````Powershell
Install-Module -Name Q-InvokeMg
````
This includes four functions that require the Microsoft.Graph Module.
````Powershell
Connect-Now
Get-QMgUser
Update-QMgUser
Update-QMgUserCSV
````
## Current functions created

### Get-QMgUser
Use Graph(Invoke-MgGraphRequest) to GET user information including: Name, Email, Department, Job title, manager and Id(Graph)

**Sample Input using -Mail:**
````powershell
Get-QMgUser -mail John.Smith@Company.com
````
*Must contain entire Email for this to work correctly.*

**Sample Output:**

    User Info

    displayName : John Smith
    mail        : John.Smith@Company.com
    jobTitle    : IT
    department  : Technology
    id          : xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
    employeeId  : 1

    Manager Info

    displayName : Mark Johnson
    mail        : Mark.Johnson@Company.com
    id          : xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
    employeeID  : 2

**Sample input using -Find:**

````powershell
Get-QMgUser -Find John
````
*Uses 'startsWith' filter parameter to locate users using displayname. Must begin with first name, as displayName cannot be filtered using 'contains' parameter*

**Sample output:**

    displayName      mail                                    id
    -----------      ----                                    --
    John Smith       John.Smith@Company.com                  xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
    John Doe         john.doe@Company.com                    xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
    John Johnson     john.Johnson@Company.com                xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

___

### Connect-Now
````powershell
Connect-Now
Please provide Access Token:
````
___

### Update-QMgUser
**Sample Input**
````powershell
Update-QMgUser -Mail John.Smith@Company.com -Parameter jobTitle -newValue 'Computer Engineer'
````
___

### Update-QMgUserCSV
**Sample Input**

````powershell
Update-QMgUserCSV -csvFile c:\temp\UserUpdate1_1_24.csv -OutFile c:\temp\UserUpdate1_1_24_Log.txt -parameterCount 1
````
*This example illustrates a csv that would update only one parameter for however many users are in the csv.*