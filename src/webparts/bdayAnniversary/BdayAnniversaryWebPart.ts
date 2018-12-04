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
}

export default class BdayAnniversaryWebPart extends BaseClientSideWebPart<IBdayAnniversaryWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.main}>
        <p class=${styles.title}>* Birthdays & Anniversaries</p>
        <ul class=${styles.content}>
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
    var today = new Date();
    var startDay = 1;
    var startMonth = 11;
    var currentYear = today.getFullYear();

    var date = new Date(); date.setDate(date.getDate() + 7); 
    var endDay = 30;
    var endMonth = 11;
    if(startMonth !== endMonth){
      this.get2months(startDay, startMonth, endDay, endMonth).then((response) => {
        this._renderList(response.value, startDay, startMonth, endDay, endMonth, currentYear, 2)
      })
    } else {
      this.get1month(startDay, startMonth, endDay)
      .then((response) => {
        this._renderList(response.value, startDay, startMonth, endDay, endMonth, currentYear, 1);
      });
    }
  }

  private get2months(startDay, startMonth, endDay, endMonth) {
    var bdayfirstMonth = this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + 
    `/_api/web/Lists/GetByTitle('Staff Events')/Items?$filter=Birth_x0020_Day gt `+ startDay + 
    ` and Birth_x0020_Month eq ` + startMonth + `'`, SPHttpClient.configurations.v1)
    .then((response) => {
      return response.json();
    });  
    var bdaySecondMonth = this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + 
    `/_api/web/Lists/GetByTitle('Staff Events')/Items?$filter=Birth_x0020_Day lt `+ endDay + 
    ` and Birth_x0020_Month eq ` + endMonth + `'`, SPHttpClient.configurations.v1)
    .then((response) => {
      return response.json();
    });
    var annifirstMonth = this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + 
      `/_api/web/Lists/GetByTitle('Staff Events')/Items?$filter=AnniversaryDay gt `+ startDay + 
      ` and AnniversaryMonth eq ` + startMonth + `'`, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      });  
      var anniSecondMonth = this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + 
      `/_api/web/Lists/GetByTitle('Staff Events')/Items?$filter=AnniversaryDay lt `+ endDay + 
      ` and AnniversaryMonth eq ` + endMonth + `'`, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      });
      return bdayfirstMonth && bdaySecondMonth && annifirstMonth && anniSecondMonth;
  }

  private get1month(startDay, startMonth, endDay){
    var bdayList = this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + 
      `/_api/web/Lists/GetByTitle('Staff Events')/Items?$filter=Birth_x0020_Day gt `+ startDay + ` and Birth_x0020_Day lt` + endDay +
      ` and Birth_x0020_Month eq ` + startMonth + `'`, SPHttpClient.configurations.v1)
        .then((response) => {
        return response.json();
      });
    var anniversaryList = this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + 
      `/_api/web/Lists/GetByTitle('Staff Events')/Items?$filter=AnniversaryDay gt `+ startDay + ` and AnniversaryDay lt ` + endDay +
      ` and AnniversaryMonth eq ` + startMonth + `'`, SPHttpClient.configurations.v1)
        .then((response) => {
        return response.json();
      });
        
      return bdayList && anniversaryList;
  }

  private _renderList(items: ISPList[], startDay, startMonth, endDay, endMonth, currentYear, numberOfMonths): void {  
    let html: string = ``;
    items.forEach((item: ISPList) => {
      html += `  
        <li>
          <div class=${styles.image }> </div>
          <div class=${styles.personWrapper}>
            <span class=${styles.name}>${item.Title}</span>
            <p class=${styles.position}>occasion</p>
            <p class=${styles.reason}>bday date</p>
            <p class=${styles.reason}>year(s)</p>
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
