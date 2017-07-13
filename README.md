# Office Mail-Addins VSTS Build/Release Task ![Build Status](https://knom.visualstudio.com/_apis/public/build/definitions/969cb309-5d7f-427d-b48a-43eeac3c4aaf/22/badge)
Install and uninstall your mail-addin to Office365 and Exchange.

## Supported Platforms ##
* Visual Studio Team Services
* Team Foundation Server 2015 Update 3 and higher ([How to install extensions in TFS](https://www.visualstudio.com/en-us/docs/marketplace/get-tfs-extensions))

## Usage ##
The extension installs the follow tasks:

![Extension Tasks](https://raw.githubusercontent.com/knom/vsts-office-tasks/master/docs/addtask.png "Extension Tasks")

* ### Install Office Outlook Add-In
    Install an Office.JS Outlook Add-In
    
    ![Screenshot](https://raw.githubusercontent.com/knom/vsts-office-tasks/master/docs/install.png "Screenshot")
    
    #### Parameters: ####
    * Path to application xml: 
        * The path of the Outlook manifest XML file. 
        * It can contain variables such as $(Build.ArtifactStagingDirectory)\manifest.xml
    * Exchange Endpoint:
        * The URL and credential to Office 365 or Exchange server. 
        * Add a new endpoint by clicking on *Manage*.
            ![Screenshot](https://raw.githubusercontent.com/knom/vsts-office-tasks/master/docs/connection1.png "Screenshot")
            ![Screenshot](https://raw.githubusercontent.com/knom/vsts-office-tasks/master/docs/connection2.png "Screenshot")
        * Use https://mail.office365.com/ as the URL for Office 365.

* ### Uninstall Office Outlook Add-In
    Remove an Office.JS Outlook Add-In from the server.
    
    #### Parameters: ####
    * Path to application xml: 
        * The path of the Outlook manifest XML file. 
        * It can contain variables such as $(Build.ArtifactStagingDirectory)\manifest.xml
    * Exchange Endpoint:
        * The URL and credential to Office 365 or Exchange server. 
        * Add a new endpoint by clicking on *Manage*.
        * Use https://mail.office365.com/ as the URL for Office 365.

		
## License ##
Published under [Apache 2.0 License](https://github.com/knom/vsts-office-tasks/blob/master/LICENSE).
