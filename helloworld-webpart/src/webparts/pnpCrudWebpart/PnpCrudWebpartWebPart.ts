import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PnpCrudWebpartWebPart.module.scss';
import * as strings from 'PnpCrudWebpartWebPartStrings';
import { getIconClassName } from '@uifabric/styling';
import { PnpHelper } from './pnphelper';
export interface IPnpCrudWebpartWebPartProps {
  listName: string;
}

export interface ISPList {
  ID: string;
  Title: string;
  Price: string;
}

export default class PnpCrudWebpartWebPart extends BaseClientSideWebPart<IPnpCrudWebpartWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
          <div class="${ styles.pnpCrudWebpart}">
              <div class="${ styles.container}">
                  <div class="${ styles.row}">
                      <div class="${ styles.column}">
                          <p class="${ styles.description}">${escape(this.properties.listName)}</p>

                          <div id="toasterMsg"></div>
                          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
                              <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                                  <input type="text" id="idPName" name="Title" placeholder="product name">
                                  <input type="text" id="idPrice" name="Price" placeholder="price">
                                  <input type="text" id="ID" name="ID" placeholder="product ID">
                                  <input type="button" class="${styles.button} create-Button" id="createBtn" value="Save Product">
                                  <input type="button" class="${styles.button} update-Button" id="updateBtn" value="Update Product">
                              </div>
                          </div>
                          <div style="background-color: white; color: black;" id="DivGetItems" />    
                          </div>  


                      </div>
                  </div>
              </div>
          </div>  
    `;
    this.setButtonsEventHandlers();
    this.getAllItems();
    document.getElementById("ID").style.visibility = "hidden";
    document.getElementById("updateBtn").style.display = "none";
  }

  private setButtonsEventHandlers(): void {
    const webPart: PnpCrudWebpartWebPart = this;
    this.domElement.querySelector('.create-Button').addEventListener('click', () => { webPart.createItem(); });
    this.domElement.querySelector('.update-Button').addEventListener('click', () => { webPart.updateItem(); });
  }

  private _getSPItems(): Promise<ISPList[]> {
    return PnpHelper.GetAllListItems(this.properties.listName).then((response) => {
      return response;
    });
  }

  private getAllItems(): void {
    this._getSPItems()
      .then((response) => {
        this._renderList(response);
      });
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';
    html += `<th>Edit</th><th>ID</th><th>Name</th><th>Price</th><th>Delete</th>`;
    if (items.length > 0) {
      items.forEach((item: ISPList) => {
        html += `    
           <tr>   
              <td> 
                <i id='${item.ID}' class="${getIconClassName('Edit')} EditLink" />
              </td>   
              <td>${item.ID}</td>    
              <td>${item.Title}</td>    
              <td>${item.Price}</td>    
              <td> 
                <i id='${item.ID}' class="${getIconClassName('Delete')} DeleteLink" />
              </td>   
           </tr>    
          `;
      });

    }
    else {
      html += "No records...";
    }
    html += `</table>`;
    const listContainer: Element = this.domElement.querySelector('#DivGetItems');
    listContainer.innerHTML = html;

    //  delete Start Bind The Event into anchor Tag
    let listItems = document.getElementsByClassName("DeleteLink");

    for (let j: number = 0; j < listItems.length; j++) {
      listItems[j].addEventListener('click', (event) => {
        let me: any = event.target;
        this.deleteItem(me.id);
      });
    }

    //  delete Start Bind The Event into anchor Tag
    let editLinks = document.getElementsByClassName("EditLink");

    for (let j: number = 0; j < editLinks.length; j++) {
      editLinks[j].addEventListener('click', (event) => {
        let me: any = event.target;
        this.editItem(me.id);
      });
    }

  }

  private editItem(id: any) {
    PnpHelper.GetListItemById(this.properties.listName, id).then(res => {
      document.getElementById('idPName')["value"] = res.Title;
      document.getElementById('idPrice')["value"] = res.Price;
      document.getElementById('ID')["value"] = res.ID;
      document.getElementById("updateBtn").style.display = "inline-block";
      document.getElementById("createBtn").style.display = "none";
    });
  }

  private deleteItem(id: number): void {
    if (!window.confirm('Are you sure you want to delete the product of id? ' + id)) {
      return;
    }

    PnpHelper.DeleteListItemById(this.properties.listName, id).then(res => {
      document.getElementById('toasterMsg').innerHTML = "Product deleted successfully";
      this.resetForm();
    });

  }

  private createItem(): void {

    const body = {
      'Title': document.getElementById('idPName')["value"],
      'Price': document.getElementById('idPrice')["value"],
    };

    PnpHelper.CreateListItem(this.properties.listName, body).then(res => {
      document.getElementById('toasterMsg').innerHTML = "Product added successfully";
      this.resetForm();
    });
  }

  private resetForm() {
    document.getElementById('idPName')["value"] = '';
    document.getElementById('idPrice')["value"] = '';
    document.getElementById('ID')["value"] = '';

    document.getElementById("updateBtn").style.display = "none";
    document.getElementById("createBtn").style.display = "inline-block";

    this.getAllItems();

  }

  private updateItem(): void {

    const itemId = document.getElementById('ID')['value'];
    const body = {
      'Title': document.getElementById('idPName')["value"],
      'Price': document.getElementById('idPrice')["value"],
    };

    PnpHelper.UpdateListItemById(this.properties.listName, itemId, body).then(res => {
      document.getElementById('toasterMsg').innerHTML = "Product updated successfully";
      this.resetForm();
    });

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
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
