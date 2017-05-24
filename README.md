# Outlook_Save_Attachments-VBA-
Save all attachments from selected emails by running this script in VBA for Outlook.

------------------------------------------------------------------------------------
'                              INSTALLATION GUIDE                                  '
------------------------------------------------------------------------------------

1. Open Outlook 2007 or newer on any machine

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
    b) Right click on "Project1" on the left hand side of the screen and select "Insert" > "Module" This will open a new tab titled "Module1"
    c) - Optional - You can change the name from "Module1" to "Save_All_Attahements" by going to "Properties" at the bottom left hand side 
        of the screen and changing the field marked "(Name)"
    d) Copy the Source from this depository and paste anywhere inside the empty space in the VBA project window.
    
Save macros as a new item on outlook

Done!
  
------------------------------------------------------------------------------------
'                                   HOW TO USE                                     '
------------------------------------------------------------------------------------

1. Highlight multiple messages on Outlook.
2. Execute the macros (either directly from the developer tab or through a custom ribbon item)
3. Select a Folder from the directory that pops up
4. Press "Yes" if you want to save the attachments to a new folder and type in the name. Be sure to include the '/' key for nested folders
5. You will be redirected to the saved folder.

NOTE: There is a small bug where you might be redirected the the "My Documents" folder but this does not affect the functionality of this macros.

------------------------------------------------------------------------------------
'                                  FINAL NOTES                                     '
------------------------------------------------------------------------------------

Feel free to comment any suggestions for improvements below!
