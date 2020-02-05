import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PublicHolidayWebPart.module.scss';
import * as strings from 'PublicHolidayWebPartStrings';
import * as microsoftTeams from "@microsoft/teams-js";
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { SPComponentLoader } from '@microsoft/sp-loader';
import 'jquery';
import 'DataTables.net';
require('bootstrap');

export interface IPublicHolidayWebPartProps {
  description: string;
}

export default class PublicHolidayWebPart extends BaseClientSideWebPart<IPublicHolidayWebPartProps> {

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

  public render(): void {
    SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.datatables.net/1.10.19/css/jquery.dataTables.min.css");

    this.domElement.innerHTML = `
      <div class="${ styles.publicHoliday}">
        <div class="${ styles.container}">
          <div id="panelLoader">
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

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
