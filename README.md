# Servicedesk-Toolbox

This project aims to be a completely portable and easily customizable tool for IT Servicedesk staff.

The user may search for e.g Active Directory users, computers or folders.
The script will then display easily customizable information for the result, and change the buttons to provide relevant tools and functions.

The script supports displaying information from Exchange. To configure this, set to enabled and to your Exchange server URI in "\Servicedesk Toolbox\configs\Exchange_config.txt"

The script supports displaying any information from Active Directory and is fully customizable by editing the template files in the folder "\Servicedesk Toolbox\configs\" through the format: "Friendly Name {AD Attribute}"

For example adding the line "Phone Number: {telephoneNumber}" to "\Servicedesk Toolbox\templates\AD_user_templates.txt":
Will display a users telephone number from Active Directory as "Phone Number: XXXXXXXXXX" when the script searches for a user.


If there are several objects the user might want to handle (e.g a user has several computers, or a folder has several AD-groups), the different objects may be picked through use of a dropdown menu in the lower right corner of the script.


The script is entirely written by me.
Feel free to use, edit, copy or take inspiration from this script in any way you'd like.


To install, simply copy the the entire folder to a relevant server in an Active Directory environment, and start through the shortcut (Servicedesk Toolbox)


![User](https://user-images.githubusercontent.com/91835664/136247783-60b4af11-bcc5-4c59-8290-6b6c3e137003.png)
![Mailbox](https://user-images.githubusercontent.com/91835664/136247812-33d66ab6-7067-42fd-aeb2-8aee78c61b25.png)
![Computer](https://user-images.githubusercontent.com/91835664/136247818-78aa548f-3c9a-4f0b-a35c-a90dd9aa9f2e.png)
![Shared Folder](https://user-images.githubusercontent.com/91835664/136247828-e94b386b-dc61-4f1d-985c-e4444762532d.png)
