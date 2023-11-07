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


export const getLargeListItems = async (context: WebPartContext, siteUrl: string, listName: string, viewName: string,  numItems: string) : Promise <any> => {
 
  const sp = spfi(siteUrl).using(SPFx(context));  

  const listView = await sp.web.lists.getByTitle(listName).views.getByTitle(viewName).select("ViewQuery")();
  const xml = `<View><Query>${listView.ViewQuery}</Query><RowLimit>${numItems}</RowLimit></View>`;
  const items = await sp.web.lists.getByTitle(listName).getItemsByCAMLQuery({ViewXml : xml});

  console.log("items", items);

  return items.map((item: any)=> {
      return {
          id: item.Id,
          title: item.Title,
          created: item.Created,
          status: item.Status, 
          empName: item.Employee_x0020_Name, 
          empNum: item.Employee_x0020_Num, 
          startDate: item.StartDate, 
          endDate: item.End_x0020_Date, 
          totalCost: item.Total_x0020_Cost, 
          approver: item.Approver, 
          totalKM: item.Total_x0020_KM, 

          // status: item.Title, // item.Status,
          // empName: item.field_1, // item.Employee_x0020_Name,
          // empNum: item.field_2, // item.Employee_x0020_Num,
          // startDate: item.field_3, // item.StartDate,
          // endDate: item.field_4, // item.End_x0020_Date,
          // totalCost: item.field_5, // item.Total_x0020_Cost,
          // approver: item.field_6, // item.Approver,
          // totalKM: item.field_7, // item.Total_x0020_KM
      }
  });

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
      // const body = JSON.stringify({Status: status});
      const body = JSON.stringify({Title: status});

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


export const updateListItem = async(context: WebPartContext, siteUrl: string,  listTitle:string, listItem: any, status: string) =>{
  const restUrl = `${siteUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${listItem.id})`;
  // const body = JSON.stringify({Status: status});
  const body = JSON.stringify({Title: status});

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