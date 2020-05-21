import { Version } from '@microsoft/sp-core-library';
import { escape } from '@microsoft/sp-lodash-subset';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'ListCrudWebPartStrings';
import styles from './ListCrudWebPart.module.scss';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IListItem } from '../models/IListItem';

export interface IListCrudWebPartProps {
  listName: string;
}

export default class ListCrudWebPart extends BaseClientSideWebPart<IListCrudWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${ styles.listCrud}">
      <div class="${ styles.container}">
          <div class="${ styles.row}">
              <div class="${ styles.column}">
                  <p class="${ styles.description}">${escape(this.properties.listName)}</p>
                  
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
                  <table class="ms-Table">
                      <thead>
                          <tr>
                              <th>ID</th>
                              <th>Product_Name</th>
                              <th>Price</th>
                              <th>Actions</th>
                          </tr>
                      </thead>
                      <tbody class="items" id="items">

                      </tbody>
                  </table>
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
    const webPart: ListCrudWebPart = this;
    this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart.createItem(); });
    this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart.updateItem(); });
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

  private createItem(): void {


    if (document.getElementById('idPName')["value"] == "") {
      alert('Required the Product Name !!!');
      return;
    }
    if (document.getElementById('idPrice')["value"] == "") {
      alert('Required Product Price !!!');
      return;
    }

    const body: string = JSON.stringify({
      'Title': document.getElementById('idPName')["value"],
      'Price': document.getElementById('idPrice')["value"],
    });

    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: body
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        return response.json();
      })
      .then((item: IListItem): void => {
        this.updateStatus(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);
      }, (error: any): void => {
        this.updateStatus('Error while creating the item: ' + error);
      });
  }

  private getAllItems(): void {

    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items?$orderby=Created desc`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      })
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((item): void => {
        var len = item.value.length;
        var txt = "";
        if (len > 0) {
          for (var i = 0; i < len; i++) {
            txt += `<tr> <td>${item.value[i].ID}</td> <td>${item.value[i].Title}</td> <td>${item.value[i].Price}</td> ` +
              `<td> <a id='${item.value[i].ID}' href='#' class='EditLink'>Edit</a>  <a id='${item.value[i].ID}' href='#' class='DeleteLink'>Delete</a></td> </tr>`;
          }
        }
        document.getElementById('items').innerHTML = txt;

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

      });

  }

  private editItem(id: any) {
    // throw new Error("Method not implemented.");
    this.getByIdItem(id).then(res => {
      document.getElementById('idPName')["value"] = res.Title;
      document.getElementById('idPrice')["value"] = res.Price;
      document.getElementById('ID')["value"] = res.ID;

      document.getElementById("updateBtn").style.display = "inline-block";
      document.getElementById("createBtn").style.display = "none";

    }, error => {
    });

  }

  private deleteItem(id: number): void {
    if (!window.confirm('Are you sure you want to delete the product of id? ' + id)) {
      return;
    }

    this.updateStatus(`Deleting item with ID: ${id}...`);
    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${id})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'
        }
      }).then((response: SPHttpClientResponse) => {
        this.updateStatus(`Item with ID: ${id} successfully deleted`);
      });
  }

  // Start Get Item By Id
  private getByIdItem(id: number): Promise<any> {


    return new Promise<string>((resolve, reject) => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse) => {
          return response.json();
        })
        .then((item): void => {
          resolve(item);
        }, (error: any): void => {
          reject(error);
        });
    });


  }
  // End Get Item By Id
  private updateStatus(status: string, items: IListItem[] = []): void {
    alert(status);
    this.getAllItems();
  }

  private updateItem() {

    let itemId = document.getElementById('ID')['value'];

    const body: string = JSON.stringify({
      'Title': document.getElementById('idPName')["value"],
      'Price': document.getElementById('idPrice')["value"],
    });

    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: body
      })
      .then((response: SPHttpClientResponse): void => {
        this.updateStatus(`Item with ID: ${itemId} successfully updated`);

        document.getElementById('idPName')["value"] = '';
        document.getElementById('idPrice')["value"] = '';
        document.getElementById('ID')["value"] = '';

        document.getElementById("updateBtn").style.display = "none";
        document.getElementById("createBtn").style.display = "inline-block";

      }, (error: any): void => {
        this.updateStatus(`Error updating item: ${error}`);
      });

  }


}
