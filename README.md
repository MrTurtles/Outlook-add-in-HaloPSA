# Introduction
This product was made for fun and educational purposes and should be maintained and updated by yourself if you decide to use it, because there is no guarantee this branch will be updated.
This product is unofficial and there was no involvement of the HaloPSA team. Be aware this could contain bugs and can be reported on the issues tab of this repository, there is no guarantuee of response. 
This product was tested in Outlook (Classic) are is not built for the Outlook Mobile app or Web app, feel free to add this functionality and don't forget to share.
This product was shared because there is no, to my knowledge, add-in made for importing email as notes in tickets and by sharing this I hope people will share there improvements for this add-in.

# Screenshots
![Outlook-ribbon-button](https://github.com/user-attachments/assets/f40d774c-c384-4a52-8c57-6f95f39a0f05)
![Login-dark-mode](https://github.com/user-attachments/assets/67da9516-0882-426a-ba23-4623777e4fce)
![select-ticket-dark-mode](https://github.com/user-attachments/assets/bce0a561-1818-4e41-a2f0-040f197b164b)
![ticket-selected-dark-mode](https://github.com/user-attachments/assets/f8013f7f-72fc-4d96-9787-5161d4eaa801)
![selected-ticket-imported-dark-mode](https://github.com/user-attachments/assets/281fd8b1-375b-480e-ba96-bcf284ff1203)

# Requirements
- HaloPSA Administrator, to create an API key with the correct permissions. Follow the tutorial for more info.
- Webserver Hosting, the files need to be hosted on a webserver where we need to point the public links in the manifest.xml
- HaloPSA Local Agent credentials, the plugin uses authentication through username and password of an Agent. If you have M365 authentication this could not work as this was not tested.

# Setup
1. Create a webserver with a public domain to host your files on. Keep it standby for the files.
2. Create an API-key in the HaloPSA portal: You do this by navigating to https://tenant.halopsa.com/config/integrations/api/applications, click New/Create, enter a name, select Username and Password as Authenticationmethod, leave the rest default and move to Permissions tab. In the Permissions tab select the following: all:teams, read:tickets, edit:tickets and read:customers, move to Security tab. In the Security tab enter the domain of the webserver you setup. After this go back to the Details tab and copy the Client-ID.
3. Follow this to get your Office Add-in ID from the manifest.xml: https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/outlook-quickstart-yo
4. Pull the files from this repository and open manifest.xml change the following placeholders: YOURADD-INID, YOURCOMPANYNAME, https://yourcompany.nl/support, https://tenant.halopsa.com. You're not required to change the assets links but it is recommended to change https://mrturtles.github.io/ to your webserver public domain. Close manifest.xml and validate it for any errors (https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest)
5. Upload your manifest.xml using aka.ms/olksideload > My Add-ins > scroll down and choose Custom > Add from file > manifest.xml. I'd recommend this now because it could take some time before the add-in shows in the Outlook Client.
6. Open taskpane.js and change the following placeholders: YOUR-HALOPSA-API-KEY, https://tenant.halopsa.com. Close all documents.
7. Upload all files to your webserver hosting and validate you can navigate to https://yourdomain.nl/taskpane.html you should see the login screen.
8. Open Outlook, navigate to an email and you should see the Add-in at the right top. Enjoy efficiency!

# Known issues
- Inline/Embedded images that are downscaled/have a low aspect ratio inside Outlook will appear as orginal size in HaloPSA when importing. If someone gets this fixed, let me know!
- Not really an issue, but the product was partly made in Dutch; for example Klant = Customer and Behandelaar = Agent.

# Contact
If you have any questions feel free to create an issues on this repository.
