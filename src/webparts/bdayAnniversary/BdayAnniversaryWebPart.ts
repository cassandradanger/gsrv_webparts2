import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {  
  SPHttpClient  
} from '@microsoft/sp-http';  

import styles from './BdayAnniversaryWebPart.module.scss';
import * as strings from 'BdayAnniversaryWebPartStrings';
import pluralize from 'pluralize';

export interface IBdayAnniversaryWebPartProps {
  description: string;
}
export interface ISPLists {
  value: ISPList[];  
}

export interface ISPList{
  Title: string;
  Employee_x0020_Birthday: string;
  Employee_x0020_Anniversary: string;
  Birth_x0020_Day: string;
  Birth_x0020_Month: string;
  AnniversaryYear: number;
  AnniversaryMonth: number;
  Email: string;
}

var today = new Date();
var currentMonth = today.getMonth() +1;
var currentYear = today.getFullYear();

var date = new Date(); date.setDate(date.getDate() + 7); 

export default class BdayAnniversaryWebPart extends BaseClientSideWebPart<IBdayAnniversaryWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.mainBA}>
        <p class=${styles.titleBA}>
          <svg class=${styles.svgBA} xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 18 18"><path fill="green" d="M9 11.3l3.71 2.7-1.42-4.36L15 7h-4.55L9 2.5 7.55 7H3l3.71 2.64L5.29 14z"/><path fill="none" d="M0 0h18v18H0z"/></svg>
          Birthdays & Anniversaries
        </p>
        <ul class=${styles.contentBA}>
          <div id="spListContainer" /></div>
        </ul>
      </div>
    </div>`;
      this._firstGetList();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }


  private _firstGetList() {
    this.context.spHttpClient.get('https://girlscoutsrv.sharepoint.com' + 
      `/_api/web/Lists/GetByTitle('Staff Events')/Items?$filter=((Birth_x0020_Month eq ` + currentMonth + `) or (AnniversaryMonth eq ` + currentMonth + '))', SPHttpClient.configurations.v1)
      .then((response)=>{
        response.json().then((data)=>{
          this._renderList(data.value)
        })
      });
    }
  
  private _renderList(items: ISPList[]): void {
    let html: string = ``;
    items.forEach((item: ISPList) => {
      let occassion = '';
      let occassionInfo = '';
      let occassion2 = '';
      let occassionInfo2 = '';
      item.Title = item.Title.toLowerCase();

      var indexOfComma = item.Title.indexOf(',');
      var firstName = item.Title.slice(indexOfComma + 2, item.Title.length);
      firstName = firstName.charAt(0).toUpperCase() + firstName.slice(1);
      var indexOfMiddleName = firstName.indexOf(' ');
      if( indexOfMiddleName !== -1){
        firstName = firstName.slice(0, indexOfMiddleName);
      }

      var lastName = item.Title.slice(0, indexOfComma);
      lastName = lastName.charAt(0).toUpperCase() + lastName.slice(1);

      var indexOf2ndLastName = lastName.indexOf(' ');
      if(indexOf2ndLastName !== -1){
        var firstLast = lastName.slice(0,indexOf2ndLastName);
        var secondLast = lastName.slice(indexOf2ndLastName, lastName.length);
        secondLast = secondLast.charAt(1).toUpperCase() + secondLast.slice(2);
        lastName = firstLast + " " + secondLast
      }
      if(item.Birth_x0020_Month === currentMonth.toString() && item.AnniversaryMonth.toString() === currentMonth.toString()){
        occassion = 'Birthday';
        occassionInfo = item.Employee_x0020_Birthday;
        occassion2 = "Anniversary";
        occassionInfo2 = pluralize('year', (currentYear - item.AnniversaryYear), true );
      } else if(item.Birth_x0020_Month === currentMonth.toString()){
        occassion = 'Birthday';
        occassionInfo = item.Employee_x0020_Birthday;
      } else if(item.AnniversaryMonth.toString() === currentMonth.toString()){
        occassion = "Anniversary";
        occassionInfo = pluralize('year', (currentYear - item.AnniversaryYear), true );
      }
      html += `  
        <li class=${styles.liBA}>
          <div class=${styles.imageBA}>
            <img class=${styles.imgBA} src="/_layouts/15/userphoto.aspx?size=L&username=${item.Email}"/>
          </div>
          <div class=${styles.personWrapperBA}>
            <span class=${styles.nameBA}>${firstName} ${lastName}</span>
            <p class=${styles.positionBA}>${occassion}</p>
            <p class=${styles.reasonBA}>${occassionInfo}</p>
            <p class=${styles.positionBA}>${occassion2}</p>
            <p class=${styles.reasonBA}>${occassionInfo2}</p>
          </div>
        </li>
        `;  
    });  
    const listContainer: Element = this.domElement.querySelector('#spListContainer');  
    listContainer.innerHTML = html;  
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
