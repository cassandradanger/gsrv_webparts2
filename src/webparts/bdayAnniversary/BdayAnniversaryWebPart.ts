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
  CurrentAnniversary: string;
  CurrentBirthDay: string;
  Employee_x0020_Anniversary: string;
  Birth_x0020_Day: string;
  Birth_x0020_Month: string;
  AnniversaryYear: number;
  AnniversaryMonth: number;
  Email: string;
  Employee: any;
}

var today = new Date();
var currentMonth =  today.getMonth() +1;
var currentYear = today.getFullYear();
var day = today.getDate();

var strToday = currentMonth + "/" + day + "/" + currentYear;

 
var datePlusSeven = new Date(); datePlusSeven.setDate(datePlusSeven.getDate() + 7); 
var monthPlus7 = datePlusSeven.getMonth() +1;
var dayPlus7 = datePlusSeven.getDate();
var yearPlus7 = datePlusSeven.getFullYear();

var strPlus7 = monthPlus7 + "/" + dayPlus7 + "/" + yearPlus7;

export default class BdayAnniversaryWebPart extends BaseClientSideWebPart<IBdayAnniversaryWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.mainBA}>
      <p class=${styles.starBA}></p>
      <p class=${styles.titleBA}>
        Birthdays & Anniversaries
      </p>
      <ul class=${styles.contentBA}>
        <div id="spListContainer" /></div>
      </ul>
    </div>`;
      this._firstGetList();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }


  private _firstGetList() {
    
    this.context.spHttpClient.get('https://girlscoutsrv.sharepoint.com' +
    `/_api/web/Lists/GetByTitle('Staff Events')/Items?$select=Title,CurrentBirthDay,CurrentAnniversary,Employee/JobTitle,Employee/SipAddress&$filter=((CurrentBirthDay ge '${strToday}') and (CurrentBirthDay le '${strPlus7}')) or ((CurrentAnniversary ge '${strToday}') and (CurrentAnniversary le '${strPlus7}'))&$expand=Employee/JobTitle,Employee/SipAddress`, SPHttpClient.configurations.v1)
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
      let anniversaryYear = item.CurrentAnniversary.slice(0, 4);
      let anniversaryMonth = item.CurrentAnniversary.slice(5, 7);
      let birthMonth = item.CurrentBirthDay.slice(5, 7);
      let birthDate = item.CurrentBirthDay.slice(8,10);
      
      if(anniversaryMonth.charAt(0) === '0'){
        anniversaryMonth = item.CurrentAnniversary.slice(6, 7);
      }
      if(birthMonth.charAt(0) === '0'){
        birthMonth = item.CurrentBirthDay.slice(6, 7);
      }
      if(birthDate.charAt(0) === '0'){
        birthDate = item.CurrentBirthDay.slice(9, 10);
      }
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
      if(birthMonth === currentMonth.toString() && anniversaryMonth.toString() === currentMonth.toString()){
        occassion = 'Birthday';
        occassionInfo =  `${birthMonth}/${birthDate}`
        occassion2 = "Anniversary";
        occassionInfo2 = pluralize('year', (currentYear - parseInt(anniversaryYear)), true );
      } else if(birthMonth === currentMonth.toString()){
        occassion = 'Birthday';
        occassionInfo = `${birthMonth}/${birthDate}`
      } else if(anniversaryMonth.toString() === currentMonth.toString()){
        occassion = "Anniversary";
        occassionInfo = pluralize('year', (currentYear - parseInt(anniversaryYear)), true );
      }
      html += `  
        <li class=${styles.liBA}>
          <div class=${styles.imageBA}>
            <img class=${styles.imgBA} src="/_layouts/15/userphoto.aspx?size=L&username=${item.Employee.SipAddress}"/>
          </div>
          <div class=${styles.personWrapperBA}>
            <span class=${styles.nameBA}>${firstName} ${lastName}</span>
            <p class=${styles.positionBA}>${item.Employee.JobTitle}</p>
            <p class=${styles.reasonBA}>${occassion}: ${occassionInfo}</p>
            <p class=${styles.reasonBA}>${occassion2} ${occassionInfo2}</p>
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
