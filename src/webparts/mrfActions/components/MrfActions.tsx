import * as React from 'react';
import styles from './MrfActions.module.scss';
import { IMrfActionsProps } from './IMrfActionsProps';
import {getLargeListItems, updateListItem} from '../../../Services/DataRequests';
import { ListView, SelectionMode} from "@pnp/spfx-controls-react/lib/ListView";
import { Stack, DefaultButton } from '@fluentui/react';

export default function MrfActions(props:IMrfActionsProps){

  const {
    hasTeamsContext,
    context
  } = props;

  console.log("props.collectionData", props.collectionData);

  const collectionData = props.collectionData ? props.collectionData : [];

  const fieldCollectionDataViewFields = collectionData.length === 0 ? [] : collectionData.map((col: any) => {
    if (col.isStatus){
      return {
          name: col.fieldName,
          displayName: col.displayName,
          sorting: col.sorting,
          minWidth: Number(col.minWidth),
          maxWidth: Number(col.maxWidth),
          render: (item: any) =>{
            switch (item[col.fieldName]){
              case 'Not Started':
                return <img width="16px" src="https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/NotStarted.png" alt='Not started' title='Not started' />;
              case 'Completed':
                return <img width="16px" src="https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/Completed.png" alt='Completed' title='Completed' />;
              case 'Deferred':
                return <img width="16px" src="https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/deferred.png" alt='Deferred' title='Deferred' />;
              case 'In Progress':
                return <img width="16px" src='https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/InProgress.png' alt='In progress' title='In progress' />;
              case 'Waiting on someone else':
                return <img width="16px" src='https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/Waiting.png' alt='Waiting on someone else' title='Waiting on someone else' />;
              case 'Exported':
                return <img width="16px" src='https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/Exported.png' alt='Exported' title='Exported' />;
              default:
                return <img width='16px' src='https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/NotStarted.png' alt='Not started' title='Not started' />
            }
          }
      }
    }
    if(col.isDate){
      return{
          name: col.fieldName,
          displayName: col.displayName,
          sorting: col.sorting,
          minWidth: Number(col.minWidth),
          maxWidth: Number(col.maxWidth),
          isResizable: col.isResizable,
          render: (item: any) =>{
            return new Date(item[col.fieldName]).toLocaleDateString();
          }
      }
    }
    if(col.isTotalCost){
      return{
        name: col.fieldName,
        displayName: col.displayName,
        sorting: col.sorting,
        minWidth: Number(col.minWidth),
        maxWidth: Number(col.maxWidth),
        isResizable: col.isResizable,
        render: (item: any) => {
          return '$' + item[col.fieldName] ;
        }
      }
    }
    if(col.isLink){
      return{
        name: col.fieldName,
        displayName: col.displayName,
        sorting: col.sorting,
        minWidth: Number(col.minWidth),
        maxWidth: Number(col.maxWidth),
        isResizable: col.isResizable,
        render: (item: any) => {
          return <a target='_blank' rel="noreferrer" data-interception="off" href={item[col.urlFieldName]}>{item[col.fieldName]}</a> ;
        }
      }
    }
    return{
      name: col.fieldName,
      displayName: col.displayName,
      sorting: col.sorting,
      minWidth: Number(col.minWidth),
      maxWidth: Number(col.maxWidth),
      isResizable: col.isResizable,
    }
  });

  const [listItems, setListItems] = React.useState([]);
  const [selItems, setSelItems] = React.useState([]);
  const [numItemsUpdated, setNumItemsUpdated] = React.useState(0);
  const [progressVis, setProgressVis] = React.useState(false);

  const percenatge = numItemsUpdated * 100/selItems.length;

  let totalCost = 0;
  if (collectionData.length !== 0){
    const totalCostColName = collectionData.filter((item: any) => item.isTotalCost)[0];
    if (totalCostColName) listItems.forEach(item => totalCost+= item[totalCostColName.fieldName]);
  }

  React.useEffect(()=>{
    if (collectionData.length !== 0){
      getLargeListItems(context, props.siteUrl, props.listName, props.viewName, props.numItems).then(res => {
        console.log("all items", res);
        setListItems(res);
      });
    }
  }, []);

  const _getSelection = (items: any[]) =>{
    console.log("selected items", items);
    setSelItems(items);
  };

  const updateItemsStatus = (status: string) => {
    const statusColName = props.collectionData.length !== 0 ? props.collectionData.filter((item: any) => item.isStatus)[0].fieldName : "Status";
    setProgressVis(true);
    const bulkUpdate = async () => {
      const updateResponseArr = [];
      for(const selItem of selItems){
        const updateResponse = await updateListItem(context, props.siteUrl, props.listName, selItem, status, statusColName);
        if (updateResponse === 1) setNumItemsUpdated(prev => prev+1);
        updateResponseArr.push(updateResponse);
      }
      Promise.all(updateResponseArr).then(()=>{
        const selIds = selItems.map(item => item.ID);
        const newListItems = listItems.map(item => {
          if (selIds.indexOf(item.ID) !== -1){
            return {...item, [statusColName]:status}
          }
          return item;
        });
        setListItems(newListItems);
      });
    };
    bulkUpdate().then(()=>{
      setProgressVis(false);
      setNumItemsUpdated(0);
    });
  };

  const viewFields = [
		{
			name: "Status",
			displayName: "Status",
			sorting: true,
			minWidth: 20,
			maxWidth: 50,
      render: (item: any) =>{
        switch (item.Status){
          case 'Not Started':
            return <img width="16px" src="https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/NotStarted.png" alt='Not started' title='Not started' />;
          case 'Completed':
            return <img width="16px" src="https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/Completed.png" alt='Completed' title='Completed' />;
          case 'Deferred':
            return <img width="16px" src="https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/deferred.png" alt='Deferred' title='Deferred' />;
          case 'In Progress':
            return <img width="16px" src='https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/InProgress.png' alt='In progress' title='In progress' />;
          case 'Waiting on someone else':
            return <img width="16px" src='https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/Waiting.png' alt='Waiting on someone else' title='Waiting on someone else' />;
          case 'Exported':
            return <img width="16px" src='https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/Exported.png' alt='Exported' title='Exported' />;
          default:
            return <img width='16px' src='https://pdsb1.sharepoint.com/sites/My-Site/SiteAssets/icons/Status/NotStarted.png' alt='Not started' title='Not started' />
        }
      }
		},
		{
			name: "Employee_x0020_Name",
			displayName: "Employee Name",
			sorting: true,
			minWidth: 100,
			maxWidth: 140,
			isResizable: true,
		},
		{
			name: "Employee_x0020_Num",
			displayName: "Employee #",
			sorting: true,
			minWidth: 100,
			maxWidth: 120,
			isResizable: true,
		},
		{
			name: "StartDate",
			displayName: "Start Date",
			sorting: true,
			minWidth: 100,
			maxWidth: 110,
			isResizable: true,
      render: (item: any) =>{
        return new Date(item.startDate).toLocaleDateString();
      }
		},
		{
			name: "End_x0020_Date",
			displayName: "End Date",
			sorting: true,
			minWidth: 100,
			maxWidth: 110,
			isResizable: true,
      render: (item: any) =>{
        return new Date(item.endDate).toLocaleDateString();
      }
		},
		{
			name: "Total_x0020_Cost",
			displayName: "Total Cost",
			sorting: true,
			minWidth: 100,
			maxWidth: 120,
			isResizable: true,
      render: (item: any) => {
        return '$' + item.totalCost ;
      }
		},
		{
			name: "Approver",
			displayName: "Approver",
			sorting: true,
			minWidth: 180,
			maxWidth: 400,
			isResizable: true,
		},
		{
			name: "Total_x0020_KM",
			displayName: "Total KM",
			sorting: true,
			minWidth: 100,
			maxWidth: 120,
			isResizable: true,
		},
		{
			name: "ID",
			displayName: "ID",
			sorting: true,
			minWidth: 100,
			maxWidth: 100,
			isResizable: true,
		},
  ];
  const WPtestJSON = [
    { 
      "fieldName": "Status",
      "displayName": "Status Test",
      "sorting": true,
      "isResizable": true,
      "minWidth": 100,
      "maxWidth": 120,
      "isDate": false,
      "isTotalCost": false,
      "isStatus": true
    },
    { 
      "fieldName": "Employee_x0020_Name",
      "displayName": "Employee Name",
      "sorting": true,
      "isResizable": true,
      "minWidth": 100,
      "maxWidth": 140,
      "isDate": false,
      "isTotalCost": false,
      "isStatus": false
    },
    { 
      "fieldName": "Employee_x0020_Num",
      "displayName": "Employee #",
      "sorting": true,
      "isResizable": true,
      "minWidth": 100,
      "maxWidth": 120,
      "isDate": false,
      "isTotalCost": false,
      "isStatus": false
    },
    { 
      "fieldName": "StartDate",
      "displayName": "Start Date",
      "sorting": true,
      "isResizable": true,
      "minWidth": 100,
      "maxWidth": 120,
      "isDate": true,
      "isTotalCost": false,
      "isStatus": false
    },
    { 
      "fieldName": "End_x0020_Date",
      "displayName": "End Date",
      "sorting": true,
      "isResizable": true,
      "minWidth": 100,
      "maxWidth": 120,
      "isDate": true,
      "isTotalCost": false,
      "isStatus": false
    },
    { 
      "fieldName": "Total_x0020_Cost",
      "displayName": "Total Cost",
      "sorting": true,
      "isResizable": true,
      "minWidth": 100,
      "maxWidth": 120,
      "isDate": false,
      "isTotalCost": true,
      "isStatus": false
    },
    { 
      "fieldName": "Total_x0020_KM",
      "displayName": "Total KM",
      "sorting": true,
      "isResizable": true,
      "minWidth": 100,
      "maxWidth": 120,
      "isDate": false,
      "isTotalCost": false,
      "isStatus": false
    },
    { 
      "fieldName": "Approver",
      "displayName": "Approver",
      "sorting": true,
      "isResizable": true,
      "minWidth": 100,
      "maxWidth": 120,
      "isDate": false,
      "isTotalCost": false,
      "isStatus": false
    },
    { 
      "fieldName": "ID",
      "displayName": "ID",
      "sorting": true,
      "isResizable": true,
      "minWidth": 100,
      "maxWidth": 100,
      "isDate": false,
      "isTotalCost": false,
      "isStatus": false
    }
  ];

  return (
    <div className={`${styles.mrfActions} ${hasTeamsContext ? styles.teams : ''}`} >

      {progressVis &&
        <>
          <div className={styles.progressBar}>
            <div className={styles.progressBarText}>Updating Items {numItemsUpdated} of {selItems.length}</div>
            <div className={styles.progressRate} style={{width: percenatge + '%'}} />
          </div>
          <div className={styles.listOverlay}/>
        </>
      }

      <div>Select Item(s) from the below list and mark them as:</div>
      <Stack className={styles.actionBtns} horizontal wrap tokens={{childrenGap: 15}}>
        <DefaultButton primary iconProps={{iconName: 'Completed12'}} onClick={()=>updateItemsStatus('Completed')}>Completed</DefaultButton> 
        <DefaultButton primary iconProps={{iconName: 'ReceiptUndelivered'}} onClick={()=>updateItemsStatus('Pending')}>Pending</DefaultButton> 
        <DefaultButton primary iconProps={{iconName: 'Clock'}} onClick={()=>updateItemsStatus('Deferred')}>Deferred</DefaultButton>        
      </Stack>

      <div className={styles.listCntnr}>
        <div className={styles.selectedItemsCount}>Selected Items: {selItems.length}</div>
        <ListView
          items={listItems}
          viewFields={fieldCollectionDataViewFields}
          compact={true}
          selectionMode={SelectionMode.multiple}
          selection={_getSelection}
          showFilter={true}
          defaultFilter=""
          filterPlaceHolder="Search..."
          dragDropFiles={true}
          stickyHeader={true}
          className={styles.listView}
        />
        <div className={styles.itemsCount}>Count: {listItems.length}</div>
        <div className={styles.totalCost}><b>Total Cost: </b>${totalCost.toLocaleString()}</div>
      </div>
    </div>
  )
}



