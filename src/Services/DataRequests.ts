/* eslint-disable @typescript-eslint/no-explicit-any */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import "@pnp/sp/items/get-all";
import {SPHttpClient, ISPHttpClientOptions} from "@microsoft/sp-http";

export const getAllViews = async (context: WebPartContext, siteUrl: string, listName: string) : Promise <any> => {
  console.log("Function getAllViews --- ");
  const sp = spfi(siteUrl).using(SPFx(context));  
  const listViews = await sp.web.lists.getByTitle(listName).views();
  return listViews;
};


export const getLargeListItems = async (context: WebPartContext, siteUrl: string, list: string, view: string,  numItems: string) : Promise <any> => {
 
  const sp = spfi(siteUrl).using(SPFx(context));  

  const listView = await sp.web.lists.getById(list).views.getById(view).select("ViewQuery")();
  const viewFields = await sp.web.lists.getById(list).views.getById(view).fields.getSchemaXml();
  
  //const xml = `<View><ViewFields>${viewFields}<FieldRef Name='FileRef' /><FieldRef Name='FileLeafRef' /></ViewFields><Query>${listView.ViewQuery}</Query><RowLimit>${numItems}</RowLimit></View>`;
  const xml = `<View><Query>${listView.ViewQuery}</Query><RowLimit>${numItems}</RowLimit></View>`;
  // const items = await sp.web.lists.getById(list).getItemsByCAMLQuery({ViewXml : xml}, 'FileRef, LinkFilename', 'File', 'Link');
  const items = await sp.web.lists.getById(list).getItemsByCAMLQuery({ViewXml : xml}, 'FileRef', 'LinkFilename', 'File', 'Link','EncodedAbsUrl','FileLeafRef','FileDirRef','LinkTitle','BaseName','_SourceUrl');

  // console.log("viewFields", viewFields);
  console.log("items before formate", items);

  return items;

  /*
  return items.map((item: any)=> {
      return {
          id: item.Id,
          title: item.Title,
          created: item.Created,
          status: item[statusCol], 
          empName: item.Employee_x0020_Name, 
          empNum: item.Employee_x0020_Num, 
          startDate: item.StartDate, 
          endDate: item.End_x0020_Date, 
          totalCost: item.Total_x0020_Cost, 
          approver: item.Approver, 
          totalKM: item.Total_x0020_KM, 
      }
  });
  */

};

/*export const isObjectEmpty = (items: any): boolean=>{
  let isEmpty:boolean = true;
  for (const item in items){
    isEmpty = items[item] === "" && isEmpty ;
  }
  return isEmpty;
};*/


export const updateListItems = async(context: WebPartContext, siteUrl: string,  listTitle:string, listItems: any, status: string) =>{
  for(const listItem of listItems){
      const restUrl = `${siteUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${listItem.id})`;
      const body = JSON.stringify({Status: status});
      //const body = JSON.stringify({Title: status});

      const spOptions: ISPHttpClientOptions = {
          headers:{
              Accept: "application/json;odata=nometadata", 
              "Content-Type": "application/json;odata=nometadata",
              "odata-version": "",
              "IF-MATCH": "*",
              "X-HTTP-Method": "MERGE",    
          },
          body: body
      };
      const _data = await context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, spOptions);
      
      if (_data.ok){
          console.log('Item(s) is/are updated!');
          return 1;
      }
  }
};


export const updateListItem = async(context: WebPartContext, siteUrl: string,  list:string, listItem: any, status: string, statusCol: string) =>{
  console.log("statusCol", statusCol);
  const restUrl = `${siteUrl}/_api/web/lists/getById('${list}')/items(${listItem.ID})`;
  const body = JSON.stringify({[statusCol]: status});
  //const body = JSON.stringify({Title: status});

  const spOptions: ISPHttpClientOptions = {
      headers:{
          Accept: "application/json;odata=nometadata", 
          "Content-Type": "application/json;odata=nometadata",
          "odata-version": "",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",    
      },
      body: body
  };
  const _data = await context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, spOptions);
  
  if (_data.ok){
      console.log('Item(s) is/are updated!');
      return 1;
  }
};