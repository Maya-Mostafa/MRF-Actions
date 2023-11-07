import * as React from 'react';
import styles from './MrfActions.module.scss';
import { IMrfActionsProps } from './IMrfActionsProps';
import {getLargeListItems, updateListItem, getAllViews} from '../../../Services/DataRequests';
import { ListView, SelectionMode} from "@pnp/spfx-controls-react/lib/ListView";
import { Stack, DefaultButton } from '@fluentui/react';

export default function MrfActions(props:IMrfActionsProps){

  const {
    hasTeamsContext,
    context
  } = props;


  const [listItems, setListItems] = React.useState([]);
  const [selItems, setSelItems] = React.useState([]);
  const [numItemsUpdated, setNumItemsUpdated] = React.useState(0);
  const [progressVis, setProgressVis] = React.useState(false);

  React.useEffect(()=>{
    console.log("React useEffect!");
    getLargeListItems(context, props.siteUrl, props.listName, props.viewName, props.numItems).then(res => {
      console.log("all items", res);
      setListItems(res);
    });

    getAllViews(context, props.siteUrl, props.listName).then(res => console.log("alll views of list " + props.listName, res));

  }, []);

  const _getSelection = (items: any[]) =>{
    console.log("selected items", items);
    setSelItems(items);
  };

  const updateItemsStatus = (status: string) => {
    
    /*updateListItems(context, props.siteUrl, props.listName, selItems, status).then(()=>{
      const selIds = selItems.map(item => item.id);
      const newListItems = listItems.map(item => {
        if (selIds.indexOf(item.id) !== -1){
          return {...item, status:status}
        }
        return item;
      });
      setListItems(newListItems);
    });*/

    setProgressVis(true);
    const bulkUpdate = async () => {
      const updateResponseArr = [];
      for(const selItem of selItems){
        const updateResponse = await updateListItem(context, props.siteUrl, props.listName, selItem, status);
        if (updateResponse === 1) setNumItemsUpdated(prev => prev+1);
        updateResponseArr.push(updateResponse);
      }
      Promise.all(updateResponseArr).then(()=>{
        const selIds = selItems.map(item => item.id);
        const newListItems = listItems.map(item => {
          if (selIds.indexOf(item.id) !== -1){
            return {...item, status:status}
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

    
    /*for(const selItem of selItems){
      updateListItem(context, props.siteUrl, props.listName, selItem, status).then(res => {
        if (res === 1) setNumItemsUpdated(prev => prev+1);
      });
    }*/

  };

  const viewFields = [
		{
			name: "status",
			displayName: "Status",
			sorting: true,
			minWidth: 20,
			maxWidth: 50,
      render: (item: any) =>{
        switch (item.status){
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
			name: "empName",
			displayName: "Employee Name",
			sorting: true,
			minWidth: 100,
			maxWidth: 140,
			isResizable: true,
		},
		{
			name: "empNum",
			displayName: "Employee #",
			sorting: true,
			minWidth: 100,
			maxWidth: 120,
			isResizable: true,
		},
		{
			name: "startDate",
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
			name: "endDate",
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
			name: "totalCost",
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
			name: "approver",
			displayName: "Approver",
			sorting: true,
			minWidth: 180,
			maxWidth: 400,
			isResizable: true,
		},
		{
			name: "totalKM",
			displayName: "Total KM",
			sorting: true,
			minWidth: 100,
			maxWidth: 120,
			isResizable: true,
		},
		{
			name: "id",
			displayName: "ID",
			sorting: true,
			minWidth: 100,
			maxWidth: 100,
			isResizable: true,
		},
  ];

  const percenatge = numItemsUpdated * 100/selItems.length;

  let totalCost = 0;
  listItems.forEach(item => totalCost+= item.totalCost);

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
          viewFields={viewFields}
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


