# SummaryAddIn

[Setup and test locally](https://github.com/freistli/SummaryAddIn/blob/main/README.md#setup)

[Build and deploy to static web host](https://github.com/freistli/SummaryAddIn/blob/main/README.md#deploy)

[Publish to organization](https://github.com/freistli/SummaryAddIn/blob/main/README.md#publish-to-organization)


## Setup
This is a quick POC to show how Outlook Add-In uses backend service to handle email content.

1. Follow the guide to install latest Node.JS, use Yeoman to create Javascript project, and use Visual Studio Code to develop:
   
   https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/outlook-quickstart?tabs=yeomangenerator

![image](https://github.com/freistli/SummaryAddIn/assets/8623897/3e5364af-abe0-40ba-a7b9-4e4a8d9613d1)

  This add in uses the same project structure, and Taskpane.js contains the main logic. 
  
  The Add-In can be tested in Classic Outlook, New Outlook, and Web Outlook.
   
2. If you are in debug phase, enable CORS for https://localhost:3000 on the service endpoint.

The icon in Web Outlook:

<img width="361" alt="image" src="https://github.com/freistli/SummaryAddIn/assets/8623897/b06c4266-f819-4348-9514-d8af5b23a407">

The Icon in Classic Outlook:

<img width="224" alt="image" src="https://github.com/freistli/SummaryAddIn/assets/8623897/3c842d0c-9645-4b99-bd1a-d68d122bd189">


The TaskPane UI:

<img width="355" alt="image" src="https://github.com/freistli/SummaryAddIn/assets/8623897/a998f7f6-3ef1-49fc-b175-2584cf4de934">

## Deploy

Here is the guide to publish the Built Script to Azure Blob Storage as a Static Web App:

https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-add-in-vs-code

Alternatively, you can also run "npm run build", and put the "dist" folder content to other web host instead of Azure Blog Static Web App. Don't forget to change localhost reference in your manifest.xml as Step 9 points out.

## Publish to Organization

Login as Global Admin or Application Admin on https://admin.microsoft.com , follow this guide to publish this AddIn manifest.xml as Integrated App in your organization, it may take up to 24 hours to show up for clients:

https://learn.microsoft.com/en-us/microsoft-365/admin/manage/office-addins?view=o365-worldwide#deploy-your-office-add-ins

After deployment, the admin site can display:

![image](https://github.com/freistli/SummaryAddIn/assets/8623897/2dfa5f83-3b7e-4410-a187-d3d8f5abf543)


Click the Integrated App item, can update or remove it:

![image](https://github.com/freistli/SummaryAddIn/assets/8623897/be7f3f9c-2f4d-40cb-8f02-513170ea61de)


About more information about the publish process and related requirement, can read:

https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish

https://learn.microsoft.com/en-us/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps?view=o365-worldwide

## Publish as Outlook Mobile Addin

To support Mobile (IOS) device, it requires to define **MobileFormFactor** in manifest.xml, the guideline is:

https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-mobile-addins

Pay attent to the [Difference](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-mobile-addins#whats-different-on-mobile) section.

Besides this, ensure the **resid** is decalred in **Resources**. 

There is a good discussion and reference in stackoverflow 

https://stackoverflow.com/questions/51641999/why-wont-this-xml-file-with-mobileformfactor-for-office-addin-work

The manifest.xml in this project is also updated to support Mobile and PC Outlook clients.