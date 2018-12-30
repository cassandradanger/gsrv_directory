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

import styles from './GsrvDirectoryWebPart.module.scss';
import * as strings from 'GsrvDirectoryWebPartStrings';

export interface IGsrvDirectoryWebPartProps {
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

export default class GsrvDirectoryWebPart extends BaseClientSideWebPart<IGsrvDirectoryWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.mainDI}>
      <ul class=${styles.contentDI}>
        <div id="spListContainer" /></div>
      </ul>
    </div>`;
      this._firstGetList();
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
      html += `  
        <li class=${styles.liDI}>
          <div class=${styles.imageDI}>
            <img class=${styles.imgDI} src="/_layouts/15/userphoto.aspx?size=L&username=${item.Email}"/>
          </div>
          <div class=${styles.personWrapperDI}>
            <span class=${styles.nameDI}>${firstName} ${lastName}</span>
            <p class=${styles.positionDI}>{position}</p>
            <p class=${styles.extensionDI}>{extension}</p>
          </div>
        </li>
        <div class=${styles.vertLineDI}></div>
        `;  
    });  
    const listContainer: Element = this.domElement.querySelector('#spListContainer');  
    listContainer.innerHTML = html;  
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
