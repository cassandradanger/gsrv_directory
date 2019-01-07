import { Version } from '@microsoft/sp-core-library';
import { sp } from "@pnp/sp";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {  
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';  

import styles from './GsrvDirectoryWebPart.module.scss';
import * as strings from 'GsrvDirectoryWebPartStrings';

export interface IGsrvDirectoryWebPartProps {
  description: string;
}
export interface ISPLists {
  value: ISPList[];
 }

export interface ISPList {
  Title:string; 
  Name:string;
  EMail:string;
  MobilePhone:string;
  Notes:string;
  SipAddress:string;
  Picture:string;
  Department:string;
  JobTitle:string;
  FirstName:string;
  LastName:string;
  WorkPhone:string;
  UserName:string;
  Id: string;
 }

 var userDept = "";

export default class GsrvDirectoryWebPart extends BaseClientSideWebPart<IGsrvDirectoryWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class=${styles.mainDI}>
      <ul class=${styles.contentDI}>
        <div id="spListContainer" /></div>
      </ul>
    </div>`;
  }

  getuser = new Promise((resolve,reject) => {
    // SharePoint PnP Rest Call to get the User Profile Properties
    return sp.profiles.myProperties.get().then(function(result) {
      var props = result.UserProfileProperties;
      var propValue = "";
      var userDepartment = "";
  
      props.forEach(function(prop) {
        //this call returns key/value pairs so we need to look for the Dept Key
        if(prop.Key == "Department"){
          // set our global var for the users Dept.
          userDept += prop.Value;
        }
      });
      return result;
    }).then((result) =>{
      this._getListData().then((response) =>{
        this._renderList(response.value);
      });
    });
  });

  public _getListData(): Promise<ISPLists> {  
    // hidden user list for the people web part
    return this.context.spHttpClient.get(`https://girlscoutsrv.sharepoint.com/_api/web/lists/getbytitle('User Information List')/items?$filter=Department eq'`+ userDept +`'`, SPHttpClient.configurations.v1)
     .then((response: SPHttpClientResponse) => {
       return response.json();
    });
  }



  
  private _renderList(items: ISPList[]): void {
    console.log(items);
    let html: string = ``;
    items.forEach((item: ISPList) => {
      let extension = item.WorkPhone ? item.WorkPhone.slice(item.WorkPhone.length - 4, item.WorkPhone.length ) : '0000';
      html += `  
        <li class=${styles.liDI}>
          <div class=${styles.imageDI}>
            <img class=${styles.imgDI} src="/_layouts/15/userphoto.aspx?size=L&username=${item.EMail}"/>
          </div>
          <div class=${styles.personWrapperDI}>
            <span class=${styles.nameDI}>${item.Title}</span>
            <p class=${styles.positionDI}>${item.JobTitle}</p>
            <p class=${styles.extensionDI}>Ext. ${extension}</p>
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
  public onInit():Promise<void> {
    return super.onInit().then (_=> {
      sp.setup({
        spfxContext:this.context
      });
    });
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
