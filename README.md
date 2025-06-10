# Build Office add-ins using Microsoft 365 Agents Toolkit
Office add-ins are integrations built by third parties into Office by using our web-based platform. This add-in template supports: Word, Excel, PowerPoint, Outlook.
Now you have the ability to create a single unit of distribution for all your Microsoft 365 extensions by using the same manifest format and schema, based on the current JSON-formatted Teams manifest.

> Note:
> The unified app manifest for Word, Excel, and PowerPoint is in preview. Visit [this link](https://aka.ms/officeversions) to check the required Office Versions. Also, publishing a unified add-in for Word, Excel, PowerPoint is not supported currently.

## Prerequisites

- [Node.js](https://nodejs.org/), supported versions: 18, 20, 22.
- Word/Excel/PowerPoint for Windows: Beta Channel, Build 18514 or higher. Outlook For Windows, Build 16425 or higher. Follow [this link](https://github.com/OfficeDev/TeamsFx/wiki/How-to-switch-Outlook-client-update-channel-and-verify-Outlook-client-build-version) for switching update channels and check your Office client build version.
- Edge installed for debugging Office add-in.
- A M365 account. If you do not have M365 account, apply one from [M365 developer program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)
- [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher.

## Debug Office add-in
- Please note that the same M365 account should be used both in Microsoft 365 Agents Toolkit and Office.
- From Visual Studio Code: Start debugging the project by choosing launch profile (default value is Word) in `Run and Debug` pane and hitting the `F5` key in Visual Studio Code. Please run VSCode as administrator if localhost loopback for Microsoft Edge Webview hasn't been enabled. Once enbaled, administrator priviledge is no longer required.

## Edit the manifest

You can find the app manifest in `./appPackage` folder. The folder contains one manifest file:
* `manifest.json`: Manifest file for Office add-in running locally or running remotely (After deployed to Azure).
You may add any extra properties or permissions you require to this file. See the [schema reference](https://raw.githubusercontent.com/OfficeDev/microsoft-teams-app-schema/preview/op/extensions/MicrosoftTeams.schema.json) for more information.

## Deploy to Azure

Deploy your project to Azure by following these steps:

| From Visual Studio Code                                                                                                                                                                                                                                                                                                                                                  | From Microsoft 365 Agents Toolkit CLI                                                                                                                                                                                                                    |
| :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| <ul><li>Open Microsoft 365 Agents Toolkit, and sign into Azure by clicking the `Sign in to Azure` under the `ACCOUNTS` section from sidebar.</li> <li>After you signed in, select a subscription under your account.</li><li>Open the Microsoft 365 Agents Toolkit and click `Provision` from LIFECYCLE section or open the command palette and select: `Microsoft 365 Agents: Provision`.</li><li>Open the Microsoft 365 Agents Toolkit and click `Deploy` or open the command palette and select: `Microsoft 365 Agents: Deploy`.</li></ul> | <ul> <li>Run command `atk auth login azure`.</li> <li>(Optional)Set environment variable AZURE_SUBSCRIPTION_ID to your subscription id in env/.env.dev or in your current shell envrionment if you are using non-interactive mode of `teamsfx` CLI.</li> <li> Run command `atk provision`.</li> <li>Run command: `atk deploy`. </li></ul> |
> Note: Provisioning and deployment may incur charges to your Azure Subscription.

To sideload the deployed add-in:

- Copy the production URL from the `ADDIN_ENDPOINT` in env/.env.dev file.
- Edit webpack.config.js file and change `urlProd` to the value you just copied. Please note to add a '/' at the end of the URL.
- Run `npm run build`.
- Run `npx office-addin-dev-settings sideload ./dist/manifest.json`.

## Validate manifest file

To check that your manifest file is valid:

- From Visual Studio Code: open the command palette and select: `Microsoft 365 Agents: Validate Application` and select `Validate using manifest schema`.
- From Microsoft 365 Agents Toolkit CLI: run command `atk validate` in your project directory.

## Known Issues
- Publish is not supported for an Office add-in project now.

# Govstream Email Assistant for Outlook

This Outlook add-in helps you generate AI-powered responses to emails using Govstream AI. It provides a simple interface to analyze the content of an email and generate appropriate responses.

## Features

- **One-click response generation**: Generate contextual responses to emails with a single click
- **AI-powered**: Uses Govstream AI to create intelligent, personalized responses
- **Simple interface**: Clean UI integrated directly in Outlook
- **Works with Outlook**: Compatible with Outlook on the web, Windows, and Mac

## How it works

1. Open an email in Outlook
2. Click the "Generate Response" button in the ribbon or open the add-in taskpane
3. The add-in will analyze the email and generate a contextual response
4. The response will be opened as a draft reply that you can review and send

## Prerequisites

- [Node.js](https://nodejs.org) (version 18 or higher)
- [npm](https://www.npmjs.com/) (included with Node.js)
- [Office 365 account](https://www.office.com/) (for testing the add-in)

## Development Setup

1. Clone this repository
2. Install dependencies with `npm install`
3. Start the development server with `npm start`

For local development and testing, use:

```bash
npm run dev-server
```

To build for production:

```bash
npm run build
```

## Deployment

The add-in can be deployed to users in your organization through the Microsoft 365 admin center or using centralized deployment.

## Configuration

The add-in connects to the Govstream API endpoint at https://testing.api.govstream.ai. To change this configuration, edit the `BACKEND_API_URL` in `src/taskpane/outlook.ts`.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For any issues or questions, please contact support at [https://govstream.ai](https://govstream.ai).