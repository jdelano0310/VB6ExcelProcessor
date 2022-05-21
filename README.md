# VB6ExcelProcessor
Uses VB6 to open - read one sheet to an array - switch to another sheet and write that data with some processing - save and close

You need MS Excel installed and referenced in your project. In VB6 with you project open, click Project in the menu then References (it is near the bottom) scroll until you find Microsoft Excel (your version) Object Libray and check the box to the left. Click OK.

This is NOT meant as a best practice or anything like that, it is meant as a super simple way of doing a few things with an Excel file, I decied to created it for a question on Reddit, being that I used VB6 for a long while (but obviously not for a while) I thought I'd give it a whirl again.

This project uses the Common Dialog Box to select the excel file via the Select button, once selected use the process button to well process the file.

My dev environment is XP SP3 Pro - 32bit in a VirtualBox (6.1.34), MS Office 2007 Pro and VB 6 version 8176

Screenshots 
When you start the app

![just started - before selecting a file](https://user-images.githubusercontent.com/8117229/169661296-78a0324b-db6a-409a-8645-cc88cbc4a418.png)

Selecting an Excel file using the common dialog box

![dialog box to select file](https://user-images.githubusercontent.com/8117229/169661299-cc8bb4b1-794c-46f8-8799-e70b636f5ee6.png)

After the file has been selected 

![after selecting a file](https://user-images.githubusercontent.com/8117229/169661304-9e652ec2-6ca2-4922-82ef-fb21878936eb.png)

After clicking the Process button the log is filled 
![after processing the selected file](https://user-images.githubusercontent.com/8117229/169661306-43683143-12d1-40ef-a583-d8b31aa4e892.png)

Note: the Excel file can not be open in Excel for it to be used.
