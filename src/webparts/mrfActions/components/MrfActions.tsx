import * as React from 'react';
import styles from './MrfActions.module.scss';
import { IMrfActionsProps } from './IMrfActionsProps';
import {getLargeListItems, updateListItem} from '../../../Services/DataRequests';
import { ListView, SelectionMode} from "@pnp/spfx-controls-react/lib/ListView";
import { Stack, DefaultButton, Spinner, MessageBar, MessageBarType, Icon, initializeIcons } from '@fluentui/react';

export default function MrfActions(props:IMrfActionsProps){ 

  console.log("All props", props);
  initializeIcons();

  const showBtns = props.showBtns === undefined ? true : props.showBtns;

  const collectionData = props.collectionData 
    ?
    props.collectionData.map((item: any) => {
      return {...item, fieldName: item.fieldName.split('|')[0], fieldType: item.fieldName.split('|')[1]};
    }) 
    : 
    [];

    //console.log("collectionData", collectionData);

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
    if (col.isApprovalIcon){
      return {
          name: col.fieldName,
          displayName: col.displayName,
          sorting: col.sorting,
          minWidth: Number(col.minWidth),
          maxWidth: Number(col.maxWidth),
          render: (item: any) =>{
            switch (item[col.fieldName]){ 
              case 'New':
              case 'Not Started':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/new.svg`)} /><span>Not Started</span></div>;
              case 'Completed':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/completed.svg`)} /><span>Completed</span></div>;
              case 'HRSpecialist_inprogress':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/submitted.svg`)} /><span>Pending HR Specialist</span></div>;
              case 'HRSpecialist_rejected':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/personRejected.svg`)} /><span>Rejected by HR Specialist</span></div>;
              case 'HRPartner_inprogress':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/submitted.svg`)} /><span>Pending HR Partner</span></div>;
              case 'HRPartner_rejected':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/personRejected.svg`)} /><span>Rejected by HR Partner</span></div>;
              case 'HRPartner_approved':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/personAccepted.svg`)} /><span>Approved by HR Partner</span></div>;
              case 'HRManager_inprogress':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/submitted.svg`)} /><span>Pending HR Manager</span></div>;
              case 'HRManager_rejected':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/personRejected.svg`)} /><span>Rejected by HR Manager</span></div>;
              case 'HRExecutive_inprogress':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/submitted.svg`)} /><span>Pending HR Executive</span></div>;
              case 'HRExecutive_rejected':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/personRejected.svg`)} /><span>Rejected by HR Executive</span></div>;
              case 'HRExecutive_approved':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/personAccepted.svg`)} /><span>Approved by HR Executive</span></div>;
              case 'Department_Accepted':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/deptAccepted.svg`)} /><span>Approved by Department</span></div>;
              case 'Department_Rejected':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/deptRejected.svg`)} /><span>Rejected by Department</span></div>;
              case 'Pending_Department_Approval':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/deptPending.svg`)} /><span>Pending Department</span></div>;
              case 'Approver1_Accepted':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/personAccepted.svg`)} /><span>Approved by Approver 1</span></div>;
              case 'Approver1_Rejected':
              case 'Employee_Rejected':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/personRejected.svg`)} /><span>Rejected by Approver 1</span></div>;
              case 'Submitted':
              case 'Pending_Employee_Approval':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/submitted.svg`)} /><span>Submitted</span></div>;
              case 'Superintendent_Accepted':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/superAccepted.svg`)} /><span>Approved by Superintendent</span></div>;
              case 'Pending_Superintendent_Approval':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/superPending.svg`)} /><span>Pending by Superintendent</span></div>;
              case 'Superintendent_Rejected':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/superRejected.svg`)} /><span>Rejected by Superintendent</span></div>;
              case 'Collecting_Feedback':
                return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/collectingFeedback.svg`)} /><span>Collecting Feedback</span></div>;
              case 'Invalid':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/invalid.svg`)} /><span>Invalid</span></div>;
              case 'Other':
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/other.svg`)} /><span>Other</span></div>;
              default:
                  return <div className={styles.formStatusCol}><img width="20" src={require(`../formIcons/other.svg`)} /><span>Other</span></div>;
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
          let itemLink = item[col.urlFieldName];
          if (col.urlFieldName === 'DisplayForm_Link') itemLink = item.EncodedAbsUrl.substring(0, item.EncodedAbsUrl.lastIndexOf('/')) + "/DispForm.aspx?ID=" + item.Id + "&ct=" + item.ContentTypeId;
          if (col.urlFieldName === 'EditForm_Link') itemLink = item.EncodedAbsUrl.substring(0, item.EncodedAbsUrl.lastIndexOf('/')) + "/EditForm.aspx?ID=" + item.Id + "&ct=" + item.ContentTypeId;
          return (
            <a target='_blank' rel="noreferrer" data-interception="off" href={itemLink}>
              {col.isEditIcon 
                ?
                <Icon iconName={'Edit'} />              
                :
                item[col.fieldName]
              }
            </a> 
          );
        }
      }
    }
    if(col.fieldName === 'Author'){
      return{
        name: col.fieldName,
        displayName: col.displayName,
        sorting: col.sorting,
        minWidth: Number(col.minWidth),
        maxWidth: Number(col.maxWidth),
        isResizable: col.isResizable,
        render: (item: any) =>{
          // console.log("item author", item);
          return item['FieldValuesAsText.Author'];
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
      render: (item:any) => {
        if (col.fieldType === 'Choice' || col.fieldType === 'Lookup'){
          return item['FieldValuesAsText.'+[col.fieldName]]
        }
        return item[col.fieldName];
      }
    }
  });

  const [listItems, setListItems] = React.useState([]);
  const [selItems, setSelItems] = React.useState([]);
  const [numItemsUpdated, setNumItemsUpdated] = React.useState(0);
  const [progressVis, setProgressVis] = React.useState(false);
  const [preloaderVisible, setPreloaderVisible] = React.useState(true);
  const [congifCollectionVisible, setConfigCollectionVisible] = React.useState(true);

  const percenatge = numItemsUpdated * 100/selItems.length;

  let totalCost = 0;
  if (collectionData.length !== 0){
    const totalCostColName = collectionData.filter((item: any) => item.isTotalCost)[0];
    if (totalCostColName) listItems.forEach(item => totalCost+= Number(item[totalCostColName.fieldName]));
  }

  React.useEffect(()=>{
    if (collectionData.length !== 0){
      setPreloaderVisible(true);
      setConfigCollectionVisible(false);
      getLargeListItems(props.context, props.siteUrl, props.list, props.view, props.numItems).then(res => {
        console.log("all items", res);
        setListItems(res);
        setPreloaderVisible(false);
      });
    }else {
      setPreloaderVisible(false);
    }

    if (props.refreshEvery5min) setInterval(refreshHandler, 300000);

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
        const updateResponse = await updateListItem(props.context, props.siteUrl, props.list, selItem, status, statusColName);
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

  const refreshHandler = () => {
    setPreloaderVisible(true);
    getLargeListItems(props.context, props.siteUrl, props.list, props.view, props.numItems).then(res => {
      console.log("reloading all items", res);
      setListItems(res);
      setPreloaderVisible(false);
    });
  };

  return (
    <div className={`${styles.mrfActions} ${props.hasTeamsContext ? styles.teams : ''}`} >

      {progressVis &&
        <>
          <div className={styles.progressBar}>
            <div className={styles.progressBarText}>Updating Items {numItemsUpdated} of {selItems.length}</div>
            <div className={styles.progressRate} style={{width: percenatge + '%'}} />
          </div>
          <div className={styles.listOverlay}/>
        </>
      }

      {congifCollectionVisible && 
        <>
        <div className={styles.welcome}>
          <img alt="" src={props.isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h3>Hi! Please configure the list information from the web part properties and the list columns by clicking on the &ldquo;Manage Data Collection&ldquo; button.</h3>
        </div>
        </>
      }

      {preloaderVisible &&
        <Spinner label="Loading data, please wait..." ariaLive="assertive" labelPosition="right" />
      }
      {!preloaderVisible && props.collectionData && props.collectionData.length !== 0 &&
        <div>
          {props.showRefresh && <a className={styles.refreshBtn} onClick={refreshHandler} href="javascript: void(0)"><Icon iconName='Refresh' />{props.refreshText}</a>}
          <div>{props.instructionText}</div>
          {showBtns &&
            <Stack className={styles.actionBtns} horizontal wrap tokens={{childrenGap: 15}}>
              <DefaultButton primary iconProps={{iconName: 'Completed12'}} onClick={()=>updateItemsStatus('Completed')}>Completed</DefaultButton> 
              <DefaultButton primary iconProps={{iconName: 'ReceiptUndelivered'}} onClick={()=>updateItemsStatus('Pending')}>Pending</DefaultButton> 
              <DefaultButton primary iconProps={{iconName: 'Clock'}} onClick={()=>updateItemsStatus('Deferred')}>Deferred</DefaultButton>        
            </Stack>
          }
          {listItems.length === 0 ?
            <MessageBar
              messageBarType={MessageBarType.warning}
              isMultiline={false}>
              Sorry, there is no data to display.
            </MessageBar>  
          :
            <div className={styles.listCntnr}>
              {props.showSelectedItemsCount && <div className={styles.selectedItemsCount}>Selected Items: {selItems.length}</div>}
                  <ListView
                    items={listItems}
                    viewFields={fieldCollectionDataViewFields}
                    compact={true}
                    selectionMode={SelectionMode.multiple}
                    selection={_getSelection}
                    showFilter={props.showFilter}
                    defaultFilter=""
                    filterPlaceHolder={props.filterPlaceholder}
                    dragDropFiles={false}
                    stickyHeader={true}
                    className={styles.listView}          
                  />
              {props.showItemsCount && <div className={styles.itemsCount}>Count: {listItems.length}</div>}
              {props.showTotalCost &&<div className={styles.totalCost}><b>Total Cost: </b>${totalCost.toLocaleString()}</div>}
            </div>
          }
        </div>
      }
    </div>
  )
}



