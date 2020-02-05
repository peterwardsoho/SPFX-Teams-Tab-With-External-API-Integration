Initial Setup of SPFx Project for Teams Tab

Follow the given steps to setup SPFx project for Teams tab as below:

1. Create a new project directory in your favorite location.

2. md public holiday

3. cd public-holiday

4. yo @microsoft/sharepoint

5. App Name: Hit Enter to have a default name (public holiday in this case) or type in any other name for your solution.

6. Target SharePoint: SharePoint Online only (Yes, I know this is for Teams, but remember the backend for Teams is SharePoint… so makes sense)

7. Place of files: We may choose to use the same folder

8. Deployment option: Selecting ‘Y’ will allow the app to be deployed instantly to all sites and will be accessible everywhere.

9. Permissions to access web APIs: Select if the components in the solution require permissions to access web APIs that are unique and not shared with other components in the tenant. Select (N)

10. Type of client-side component to create: We can choose to create a client-side web part or an extension. Selected WebPart

11. Web Part Name: PublicHoliday

12. Web part description: Enter the description but it is not mandatory.

13. Framework to use: Select No JavaScript Framework
![](https://github.com/peterwardsoho/SPFX-Teams-Tab-With-External-API-Integration/blob/master/FrameworkToUse.png)
14.	At this point, Yeoman installs the required dependencies and scaffolds the solution files. Creation of the solution might take a few minutes. Yeoman scaffolds the project to include your PublicHoliday web part as well.
15.	Next, open the Public Holiday SPFx project in Visual Studio Code.
16.	Updating the web part manifest to make it available for Microsoft Teams.
Locate the manifest json file for the web part you want to make available to Teams and modify the supportedHosts properties to include "TeamsTab" as in the following example.
![](https://github.com/peterwardsoho/SPFX-Teams-Tab-With-External-API-Integration/blob/master/SupportedHosts.png)
17.	Add the below code in the webpart.ts files
•	import * as microsoftTeams from "@microsoft/teams-js";
•	In a class, define a variable to store Microsoft Teams context.
•	private _teamsContext: microsoftTeams.Context;
•	Add the onInit() method to set the Microsoft Teams context.

private _teamsContext: microsoftTeams.Context;
  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }
Implement External API

1.	Add the below code in the webpart.ts file
•	import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
private _getPublicHolidayFromExternalApi(): Promise<any> {
    return this.context.httpClient
      .get(
        'https://calendarific.com/api/v2/holidays?&api_key=69f7fb86d8de5f3121b5e6783b7b4fb1d2424ba7&country=us&year=' + new Date().getFullYear(),
        HttpClient.configurations.v1
      )
      .then((response: HttpClientResponse) => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse;
      }) as Promise<any>;
  }
![](https://github.com/peterwardsoho/SPFX-Teams-Tab-With-External-API-Integration/blob/master/HTTPClient.png)

