import * as pnp from 'sp-pnp-js';
import { ItemAddResult, ItemUpdateResult } from 'sp-pnp-js';

export class PnpHelper {


    /**
     * Get all items from sharepoint List
     */
    public static GetAllListItems(listName: string): Promise<any> {
        return pnp.sp.web.lists.getByTitle(listName).items.get();
    }

    /**
    * Get item from sharepoint List by id
    */
    public static GetListItemById(listName: string, id: number): Promise<any> {
        return pnp.sp.web.lists.getByTitle(listName).items.getById(id).get();
    }

    /**
     * Create an entry into the SharePoint List
     */
    public static CreateListItem(listName: string, body: any): Promise<ItemAddResult> {
        return pnp.sp.web.lists.getByTitle(listName).items.add(body);
    }

    /**
     * Update SharePoint List Item by Id
     */
    public static UpdateListItemById(listName: string, id: number, body: any): Promise<ItemUpdateResult> {
        return pnp.sp.web.lists.getByTitle(listName).items.getById(id).update(body);
    }

    /**
     * Delete SharePoint List Item by Id
     */
    public static DeleteListItemById(listName: string, id: number): Promise<void> {
        return pnp.sp.web.lists.getByTitle(listName).items.getById(id).delete();
    }

}

