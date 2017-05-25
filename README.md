# Outlook_Save_Attachments-VBA-
Save all attachments from selected emails by running this script in VBA for Outlook.

------------------------------------------------------------------------------------
'                               INSTALLATION GUIDE                                 ' 
------------------------------------------------------------------------------------

Save main script as a (.bas) file
    
    a) Copy script and paste it in notepad (will be installed by default on any 'Windows' computer.
    b) Click on "File" then "Save As". A new window should pop up.
    c) Under the drop down that says "Save as type" select All Files.
    d) Type "Save_all_attachments.bas" in the file name box and click save. 
    NOTE: Be sure to note where you saved this file as you will beed it later

Open Outlook 2007 or newer on any machine

Enable macros

    a) Go to "File" on the top left hand corner of the window
    b) Select the "Option" tab on the left hand side (this may be different for outlook 2007 or older. A new window should pop up.
    c) Select "Trust Center" on the bottom left hand side of the new window
    d) Select the "Trust Center Settings..." button on the right side of the window. A new window should pop up.
    e) Select "Macro Settings" on the left hand side of the window.
    f) It is highly recommended that you select "Notifications for all macros" to prevent any security breaches but selecting 
      "Enable all macros" should work too.

Create a new module and paste the source code within

    a) Press the ALT-F11 keys on your keyboard. the VBA Project Editor window should pop up
    b) Drag and drop the source code anywhere inside the VBA Project Editor
    
Save macros as a new item on outlook
    
    a) Go to "File" on the top left hand corner of the window
    b) Select the "Option" tab on the left hand side (this may be different for outlook 2007 or older. A new window should pop up.
    c) Select "Customize Ribbon on the left Pane.
    d) Select "Macros" on the left drop-down menu and select "Project1.Project1.Save_all_attachments"
    e) Finally select where you want to add your new custom too in the ribbon and press the "Add >>" button.
    

Done!
  
------------------------------------------------------------------------------------
'                                  HOW TO USE                                      '
------------------------------------------------------------------------------------

1. Highlight multiple messages on Outlook.
2. Execute the macros (either directly from the developer tab or through a custom ribbon item)
3. Select a Folder from the directory that pops up
4. Press "Yes" if you want to save the attachments to a new folder and type in the name. Be sure to include the '/' key for nested folders
5. You will be redirected to the saved folder.

NOTE: There is a small bug where you might be redirected the the "My Documents" folder but this does not affect the functionality of this macros.

------------------------------------------------------------------------------------
'                                  THE MAKING OF                                   '  
------------------------------------------------------------------------------------

The Problem
------------------------------------------------------------------------------------
During my time at work, I was tasked to upload hardcopies of old documents onto our server. This was fairly manual work. I would take huge binders of documents, remove each one and scan them. The problem was when I got back to my inbox to find several hundred emails from our scanner. I’m not going to just sit there, open each email, save the attachments and move on to the next. I’m going to write a macro to take care of that for me.

Summary
------------------------------------------------------------------------------------
This project was an interesting one as it combined both Outlook’s VBA library and VB.NET to grab attachments from multiple emails and save them to a specific folder on the desktop. Later I was able to add functionality by customising which folder the end-user could save to and even create a new directory if the one provided didn’t exist. Although this was just a simple tool, I found it to be extremely productive when scanning large volumes of documents.

How it works
------------------------------------------------------------------------------------
When the macro is initiated, it will prompt the end-user with directory explorer GUI. Once the user selected a directory, the function will convert the address of the selected directory into a string and closing it with a backslash. The user will then be prompted whether they want to create a new folder with the option to nest multiple folders by placing a backslash in between nested folders. A while loop with then cycle through the entered string and create a new directory for every folder that does not exist. Once a new string has been generated, the macro will then cycle through each attachment through each email and manually save the file by concatenate the directory string with the name of the attachment through a shell.

On a Side Note
------------------------------------------------------------------------------------
This project was also quite terrifying as it gave me the paranoid possibility that this script could theoretically save the attachments of every email in my inbox to an unknown directory. Luckily this did not happen but just to be safe, I created a dummy outlook account as to limit the impact of a bug gone wrong.

