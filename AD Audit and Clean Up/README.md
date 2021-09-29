A script that collects data from Active Directory, Azure Active Directory, Exchange, Skype for Business, and Intune. It collects some of that data in Azure Storage Tables so you can use it for reporting, disables and/or deletes objects that are stale and emails you a report. 

**WARNING WARNING WARNING WARNING WARNING WARNING WARNING** 

This script is designed to delete objects from your Active Directory, Azure Active Directory, Exchange, Intune and SQL. Use at your own risk. 
I provide no warranty or support and accept no responsibility. You should never run random scripts off the internet without understanding their content. 

Sample Azure Storage Table for AD Users
![image](https://user-images.githubusercontent.com/55263241/135212516-b97e9508-5b19-453a-8cd8-3ce2dbb5597a.png)

Sample Azure Storage Table for AAD Computers. The amount of device data you can capture depends on enrolment type, Intune, etc. 
![image](https://user-images.githubusercontent.com/55263241/135212880-b26bdd09-ed69-4b25-9667-cd894913bd56.png)

You can create Power BI dashboards from the Azure Storage Table data, help you identify trends like volumes of stale accounts, etc
![image](https://user-images.githubusercontent.com/55263241/135213228-ad69d9fb-4574-4031-99b3-3758926b0205.png)

When accounts are disabled or deleted when you run this script, you can get a dynamic HTML report emailed to you. 
![image](https://user-images.githubusercontent.com/55263241/135213649-22ad602a-385b-4ebd-8664-55b0c42e1996.png)

Many thanks to https://github.com/EvotecIT for the wonderful module, and anyone else I have borrowed from. 
