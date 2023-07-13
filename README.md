# Outlook-Mail-Notes

This Add-In allows you to add notes to e-mails in Outlook.
The notes are synced with your Exchange account, which allows you to access them from anywhere you can access your e-mail.

## Installation

### Outlook Desktop client

- Use *Home* > *Get Add-Ins* to navigate to the *Add-Ins* page (note that Add-Ins can only be installed in mail clients that are connected to an Exchange server, i.e. IMAP accounts cannot use Add-Ins).
- Go to the *My Add-Ins* tab, scroll to the bottom and click on the *+ Add a custom add-in* button.
	- To always automatically get the latest version of the Add-In, select *Add from URL* and enter the following URL: `https://nikkelm.github.io/Outlook-Mail-Notes/manifest.xml`
	- Alternatively, follow the instructions in the [Manual Installation](#manual-installation) section below to install the Add-In from a local file.
- The Add-In will now be available in the *Home* tab of your Outlook client. You can use it by selecting an e-mail and clicking on the *Take notes* button.

### Outlook for the Web

- Using Outlook for the Web, click on the gear icon in the top right corner and select *Manage Add-Ins*.
- Use the *+* button to add a new Add-In.
	- To always automatically get the latest version of the Add-In, select *Add from URL* and enter the following URL: `https://nikkelm.github.io/Outlook-Mail-Notes/manifest.xml`
	- Alternatively, follow the instructions in the [Manual Installation](#manual-installation) section below to install the Add-In from a local file.
- The Add-In will now be available when viewing an e-mail, its icon will be visible in the top right corner of the e-mail.

### Manual Installation

- Clone the repository or download a relevant release and navigate to the project folder.
- Run `npm install` to install all required dependencies.
- Run `npm run build` to build the project. All relevant files will be placed in the `dist` folder.
- Follow the first steps of the relevant [Installation](#installation) section above, but select *Add from file* instead of *Add from URL*.
- Select the `manifest.xml` file from the `dist` folder. *Note that by installing the Add-In from your local machine will prevent you from receiving any relevant updates.*

## Contribute

- Clone the repository and navigate to the project folder.
- Run `npm install` to install all required dependencies.
- To start a development server, use `npm run watch`. To also automatically register the Add-In in your Outlook desktop client, use `npm start` instead.

----

Do you have any feedback or questions? Feel free to open an [issue](https://github.com/NikkelM/Outlook-Mail-Notes/issues).