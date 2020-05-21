import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PnpCrudWebpartWebPart.module.scss';
import * as strings from 'PnpCrudWebpartWebPartStrings';
import * as pnp from 'sp-pnp-js';
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
                                  <button class="${styles.button} create-Button" id="createBtn">
                                      <span class="${styles.label}">Save Product</span>
                                  </button>
                                  <button class="${styles.button} update-Button" id="updateBtn">
                                      <span class="${styles.label}">Update Product</span>
                                  </button>
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
    this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart.createItem(); });
    this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart.updateItem(); });
  }

  private _getSPItems(): Promise<ISPList[]> {
    return pnp.sp.web.lists.getByTitle(this.properties.listName).items.get().then((response) => {
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
    let html: string = '<table class="TFtable" border=1 width=style="bordercollapse: collapse;">';
    html += `<th></th><th>ID</th><th>Name</th><th>Price</th><th>Actions</th>`;
    if (items.length > 0) {
      items.forEach((item: ISPList) => {
        html += `    
           <tr>   
              <td>  <input type="radio" id="ProductId" name="ProductId" value="${item.ID}"> <br> </td>   
              <td>${item.ID}</td>    
              <td>${item.Title}</td>    
              <td>${item.Price}</td>    
              <td> <a id='${item.ID}' href='#' class='EditLink'>Edit</a>  <a id='${item.ID}' href='#' class='DeleteLink'>Delete</a></td>   
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
    debugger;
    pnp.sp.web.lists.getByTitle(this.properties.listName).items.getById(id).get().then(res => {
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

  }

  private createItem(): void {
    pnp.sp.web.lists.getByTitle(this.properties.listName).items.add({
      Title: document.getElementById('idPName')["value"],
      Price: document.getElementById('idPrice')["value"],
    }).then(res => {
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
    pnp.sp.web.lists.getByTitle(this.properties.listName).items.getById(itemId).update({
      Title: document.getElementById('idPName')["value"],
      Price: document.getElementById('idPrice')["value"]
    }).then(res => {
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
