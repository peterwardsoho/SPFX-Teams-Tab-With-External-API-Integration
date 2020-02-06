**Initial Setup of SPFx Project for Teams Tab**


1.	Create a new project directory in your favorite location. 
2.	md public holiday
3.	cd public-holiday
4.	yo @microsoft/sharepoint
5.	App Name: Hit Enter to have a default name (public holiday in this case) or type in any other name for your solution.
6.	Target SharePoint: SharePoint Online only (Yes, I know this is for Teams, but remember the backend for Teams is SharePoint… so makes sense)
7.	Place of files: We may choose to use the same folder
8.	Deployment option: Selecting ‘Y’ will allow the app to be deployed instantly to all sites and will be accessible everywhere.
9.	Permissions to access web APIs: Select if the components in the solution require permissions to access web APIs that are unique and not shared with other components in the tenant. Select (N)
10.	Type of client-side component to create: We can choose to create a client-side web part or an extension. Selected WebPart
11.	Web Part Name: PublicHoliday
12.	Web part description: Enter the description but it is not mandatory.
13.	Framework to use: Select No JavaScript Framework
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


**Implement External API**

1.	Add the below code in the webpart.ts file
```
	import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
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
```
**Implement jQuery, Datatable and Bootstrap for Design**

1.	Install below npm packages. <br/>
•	npm install @types/jquery@2 --save <br/>
•	npm install @types/jqueryui --save <br/>
•	npm install datatables.net <br/>
•	npm install datatables.net-jqui <br/>
•	npm install --save @types/datatables.net

2.	Navigate to Config.json file under Config Folder. (Config > Config.json) and add the below code in external node.

```
"jquery": {
      "path": "https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js",
      "globalName": "jquery"
    },
    "bootstrap": {
      "path": "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js",
      "globalName": "bootstrap",
      "globalDependencies": ["jquery"]
    }
```
![](https://github.com/peterwardsoho/SPFX-Teams-Tab-With-External-API-Integration/blob/master/ImplementJQueryConfig.png)

3.	Add the below code in the webpart.ts file.
•	import { SPComponentLoader } from '@microsoft/sp-loader';
•	import 'jquery';
•	import 'DataTables.net';
•	require('bootstrap');
•	Update render() method as below.

```
public render(): void {
    SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.datatables.net/1.10.19/css/jquery.dataTables.min.css");

    this.domElement.innerHTML = `
      <div class="${ styles.publicHoliday}">
        <div class="${ styles.container}">
          <div>
            <div class="panel panel-primary">
              <div class="panel-heading">
                <h4 class="panel-title">
                  Public Holidays
                  </h4>
              </div>
              <div class="panel-body">
                <table id="tblPublicHolidays" class=${styles.publicHolidayTable} width="100%">
                </table>
              </div>
            </div>
          </div>
        </div>
      </div>`;
      this.RenderPublicHolidayEntries();
  }
  ```

•	Add the below RenderPublicHolidayEntries() method in webpart.ts file.
```
private RenderPublicHolidayEntries(): void {
    this.context.statusRenderer.displayLoadingIndicator(document.getElementById("tblPublicHolidays"),"Please wait...");
    var holidaysData = [];
    var holidaysDataJsonArr = [];
    this._getPublicHolidayFromExternalApi().then((items) => {
      holidaysData = items.response.holidays;
      for (var i = 0; i < holidaysData.length; i++) {

        let states: any;
        if (holidaysData[i].states == "All") {
          states = holidaysData[i].states;
        }
        else {
          var statesArr = [];
          for (var j = 0; j < holidaysData[i].states.length; j++) {
            statesArr.push(holidaysData[i].states[j].name);
          }
          states = statesArr.toString();
        }
        holidaysDataJsonArr.push({ Holiday: holidaysData[i].name, Description: holidaysData[i].description, Date: new Date(holidaysData[i].date.iso).toLocaleDateString(), HolidayType: holidaysData[i].type.toString(), Locations: holidaysData[i].locations, States: states });
      }

      var jsonArray = holidaysDataJsonArr.map(function (item) {
        return [
          item.Holiday,
          item.Description,
          item.Date,
          item.HolidayType,
          item.Locations,
          item.States
        ];
      });

      $('#tblPublicHolidays').DataTable({
        data: jsonArray,
        columns: [
          { title: "Holiday" },
          { title: "Description" },
          { title: "Date" },
          { title: "Holiday Type" },
          { title: "Locations" },
          { title: "States" }

           ]
      });
      this.context.statusRenderer.clearLoadingIndicator(document.getElementById("tblPublicHolidays"));
    }).catch((error) => {
      console.log("Something went wrong " + error);
      this.context.statusRenderer.clearLoadingIndicator(document.getElementById("tblPublicHolidays"));
    });
  }
  ```
  

•	Add the below style for table in WebPart.module.scss file.

```
.container {
    width: 100%;
    margin: 0px auto;
    box-shadow: 0 2px 4px 0 rgba(0, 0, 0, 0.2), 0 25px 50px 0 rgba(0, 0, 0, 0.1);
    border-radius: 4px;
  }
```

```
/*Table style*/

  .publicHolidayTable {
    table-layout: fixed;
    width: 100%;
  }
    background-color: #337ab7;
    color: #fff;
  }

  .publicHolidayTable th, .publicHolidayTable td{
    border: 1px solid #d1cfd0;
    text-align: center;
    word-break: break-all;
    padding: 5px;
  }

  .publicHolidayTable tr.table-heading{
      background-color: #337ab7;
      font-size: 18px;
      font-weight: bold;
  }
.publicHolidayTable .table-links{
      color: #058ac8;
  }
  
  .publicHolidayTable .table-links:hover{
    color: #5e9732;
  }
  
    /*Table style*/

```

**Deployment Process**

1.	Bundle the solution run gulp bundle --ship.
2.	Package the solution run gulp package-solution –ship
3.	Navigate to the solution and copy public-holiday.sppkg file from SharePoint folder.
4.	Add this public-holiday.sppkg into the Appcatlog of your tenant.
Use the App Catalog to make custom business apps available for your SharePoint Online environment.
5.	Make sure you Check “Make this solution available to all sites in the organization” option and click on “Deploy”.
![](https://github.com/peterwardsoho/SPFX-Teams-Tab-With-External-API-Integration/blob/master/DeploymentProcess.png)
6.	Once you deploy this public-holiday.sppkg file.
7.	Click “Sync to Teams” in files tab from the ribbon as shown below.
![](https://github.com/peterwardsoho/SPFX-Teams-Tab-With-External-API-Integration/blob/master/SyncToTeams.png)
8.	Once it has done open your MS Teams and navigate to the team where you wanted to add this teams tab.
9.	Click on the + icon and a dialog will open with Add a tab title
![](https://github.com/peterwardsoho/SPFX-Teams-Tab-With-External-API-Integration/blob/master/AddTabTop.png)
![](https://github.com/peterwardsoho/SPFX-Teams-Tab-With-External-API-Integration/blob/master/AddATab.png)
10.	Select the Public Holiday WebPart with your title from the list of app has and click it. It will open a new dialog box as shown below and click Save to add it.
![](https://github.com/peterwardsoho/SPFX-Teams-Tab-With-External-API-Integration/blob/master/PublicHolidaySave.png)
11. Final output of Public Holiday teams tab will be as below.
![](https://github.com/peterwardsoho/SPFX-Teams-Tab-With-External-API-Integration/blob/master/FinalOutput.png)