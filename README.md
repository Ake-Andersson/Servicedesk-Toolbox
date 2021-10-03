# Servicedesk-Toolbox

This project aims to be a completely portable and easily customizable tool for IT Servicedesk staff.

The user may search for e.g Active Directory users, computers or folders.
The script will then display easily customizable information for the result, and change the buttons to provide relevant tools and functions.


The script supports displaying any information from Active Directory and is fully customizable by editing the template files in the folder "\Servicedesk Toolbox\templates\" through the format: "Friendly Name {AD Attribute}"

For example adding the line "Phone Number: {telephoneNumber}" to "\Servicedesk Toolbox\templates\AD_user_templates.txt":
Will display a users telephone number from Active Directory as "Phone Number: XXXXXXXXXX" when the script searches for a user.


If there are several objects the user might want to handle (e.g a user has several computers, or a folder has several AD-groups), the different objects may be picked through use of a dropdown menu in the lower right corner of the script.


The script is entirely written by me.
Feel free to use, edit, copy or take inspiration from this script in any way you'd like.


To install, simply copy the the entire folder to a relevant server in an Active Directory environment, and start through the shortcut (Servicedesk Toolbox)


![User](https://user-images.githubusercontent.com/91835664/135761939-d5771494-e8a0-4674-85c1-773ae3e584ee.PNG)
![Computer](https://user-images.githubusercontent.com/91835664/135762523-cac4e021-e5c0-469c-bb17-6f3a4ba06c53.PNG)
![Shared Folder](https://user-images.githubusercontent.com/91835664/135763001-0704bf85-3a24-472a-b50a-bb30412c7114.PNG)
