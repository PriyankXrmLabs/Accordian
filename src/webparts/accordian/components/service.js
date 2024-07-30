import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { DescriptionFieldLabel } from "AccordianWebPartStrings";





export async function addToList(context , listName, title , desc) {
    
      console.log(title,desc)
     const sp = spfi().using(SPFx(context));

     try {
      const item = await sp.web.lists.getByTitle(listName).items.add({
        Title:title,
        Description:desc,
      });
      
      console.log(item);
      } catch (error) {
        console.error("Error adding item: ", error);
      }



}
export async function getData(context , listName) {
    
 
     const sp = spfi().using(SPFx(context));

     try {
      const item = await sp.web.lists.getByTitle(listName).items();
        
      console.log(item);
      return item;
      } catch (error) {
        console.error("Error adding item: ", error);
      }



}