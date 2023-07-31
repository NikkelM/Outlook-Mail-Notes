<h1 align="center">Outlook Mail Notes</h1>

<p align="center">
<a href="https://github.com/NikkelM/Outlook-Mail-Notes/tree/main/CHANGELOG.md">
  <img src="https://img.shields.io/badge/view-changelog-blue"
    alt="View changelog"></a>
</p>

This Add-In allows you to add notes to e-mails in Outlook.
The notes are synced with your Exchange account, which allows you to access them from anywhere you can access your e-mail.

## Installation

### Outlook Desktop client

- Use _Home_ > _Get Add-Ins_ to navigate to the _Add-Ins_ page (note that Add-Ins can only be installed in mail clients that are connected to an Exchange server, i.e. IMAP accounts cannot use Add-Ins).
- Go to the _My Add-Ins_ tab, scroll to the bottom and click on the _+ Add a custom add-in_ button.
- Select _Add from URL_ and enter the following URL: `https://nikkelm.dev/Outlook-Mail-Notes/manifest.xml`
- The Add-In will now be available in the _Home_ tab of your Outlook client. You can use it by selecting an e-mail and clicking on the _Take notes_ button.

### Outlook for the Web

- Using Outlook for the Web, click on the gear icon in the top right corner and select _Manage Add-Ins_.
- Use the _+_ button to add a new Add-In.
- Select _Add from URL_ and enter the following URL: `https://nikkelm.dev/Outlook-Mail-Notes/manifest.xml`
- The Add-In will now be available when viewing an e-mail, its icon will be visible in the top right corner of the e-mail.

## Updates

_TL;DR: Updates to the Add-In's main functionality will be automatically applied, but updates to the Add-In's `manifest.xml` file require a re-installation due to how Add-In sideloading works._

Office Add-Ins are hosted on the developer's server - in the case of this Add-In, this is `https://nikkelm.dev`.
All files, such as the displayed note editor and relevant background logic are hosted on this server, and not downloaded to your machine.
This means that any released updates will automatically reflect in your client.

Please note that while you will automatically receive any updates to the files the Add-In uses in the background, Outlook will **not** automatically update the Add-In's `manifest.xml`.
This means that in the case that anything in this file is changed, this will not reflect in your client until you manually remove and re-add the Add-In.
Unless the Add-In breaks, and until I am able to add an option to export/import notes, I highly discourage you from doing this, as all of your notes will be lost.
Changes to the `manifest.xml` file will not happen frequently, as doing so is only necessary when adding new ways to interact with the Add-In from within the Outlook client.

In any case, always feel free to open an [issue](https://github.com/NikkelM/Outlook-Mail-Notes/issues) if you have any questions or concerns.

## Contribute

- Clone the repository and navigate to the project folder.
- Run `npm install` to install all required dependencies.
- To start a development server, use `npm run watch`. To also automatically register the Add-In in your Outlook desktop client, use `npm start` instead.

---

Do you have any feedback or questions? Feel free to open an [issue](https://github.com/NikkelM/Outlook-Mail-Notes/issues).

If you enjoy this Add-In and want to say thanks, consider buying me a [coffee](https://ko-fi.com/nikkelm) or [sponsoring](https://github.com/sponsors/NikkelM) this project.
