import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'MrfActionsWebPartStrings';
import MrfActions from './components/MrfActions';
import { IMrfActionsProps } from './components/IMrfActionsProps';

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/fields";

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { PropertyFieldSitePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldSitePicker';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldViewPicker, PropertyFieldViewPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldViewPicker';

export interface IMrfActionsWebPartProps {
  description: string;
  numItems: string;
  siteUrl: string;
  collectionData: any[];

  list: string;
  view: string;

  instructionText: string;
  showFilter: boolean;
  filterPlaceholder: string;
  showSelectedItemsCount: boolean;
  showItemsCount: boolean;
  showTotalCost: boolean;
}

export default class MrfActionsWebPart extends BaseClientSideWebPart<IMrfActionsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IMrfActionsProps> = React.createElement(
      MrfActions,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,

        context: this.context,
        numItems: this.properties.numItems,
        siteUrl: this.properties.siteUrl,
        collectionData: this.properties.collectionData,
        
        list: this.properties.list,
        view: this.properties.view,

        instructionText: this.properties.instructionText,
        showFilter: this.properties.showFilter,
        filterPlaceholder: this.properties.filterPlaceholder,
        showSelectedItemsCount: this.properties.showSelectedItemsCount,
        showItemsCount: this.properties.showItemsCount,
        showTotalCost: this.properties.showTotalCost,

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // return this._getEnvironmentMessage().then(message => {
    //   this._environmentMessage = message;
    // });
    console.log("onInit----");
    return super.onInit();
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private colInternalName: any;
  private async loadListColumns() : Promise<any>{
    console.log("loadListColumns Fnc");
    try {
      const sp = spfi(this.properties.siteUrl).using(SPFx(this.context));  
      const listFields: any = await sp.web.lists.getById(this.properties.list).fields();

      console.log("all list fields", listFields);

      // return listFields.filter((col: any)=>{
      //   const excludedCols = ['_CheckinComment', '_ColorTag', '_CommentCount', '_ComplianceFlags', '_ComplianceTag', '_ComplianceTagUserId', '_ComplianceTagWrittenTime', '_CopySource', '_DisplayName', '_IsRecord', '_LikeCount', '_UIVersionString', 'Edit', 'Open','DocIcon','FileSizeDisplay','AppAuthor','AppEditor','ComplianceAssetId','ContentType','FolderChildCount','ItemChildCount','LinkFilenameNoMenu','MediaServiceImageTags','ParentLeafName','ParentVersionString'];
      //   if (col.Hidden || excludedCols.indexOf(col.InternalName) !== -1) return false;
      //   return true;
      // })

      listFields.push({EntityPropertyName: 'DisplayForm_Link'}, {EntityPropertyName: 'EditForm_Link'});

      return listFields.sort((a:any, b: any) => a.EntityPropertyName.localeCompare(b.EntityPropertyName)).map((col:any)=>{
        return {
          key: col.EntityPropertyName,
          text: col.EntityPropertyName
        }
      });

    }catch(error){
      console.log("error", error);
    }
  }

  /* Loading Dpd with list view names - Start */
  // private views: IPropertyPaneDropdownOption[];
  // private async loadListViews(): Promise<IPropertyPaneDropdownOption[]> {    
  //   console.log("loadListViews() --- ");
  //   const viewsTitle : any = [];
  //   try {
  //     const sp = spfi(this.properties.siteUrl).using(SPFx(this.context));  
  //     const listViews = await sp.web.lists.getByTitle(this.properties.listName).views();

  //     if (listViews) {
  //       listViews.map((result: any)=>{
  //         if (result.Title !== ''){
  //           viewsTitle.push({
  //             key: result.Title,
  //             text: result.Title
  //           });
  //         }
  //       });
  //       return viewsTitle;
  //     }
  //   } catch (error) {
  //     console.log("error", error);
  //   }
  // }
  protected onPropertyPaneConfigurationStart(): void {
    console.log("onPropertyPaneConfigurationStart() -------------- ");
    // if (this.views) {
    //   this.render();  
    //   return;
    // }
    /*this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'views');
    this.loadListViews().then((viewOptions: IPropertyPaneDropdownOption[]): void => {
        this.views = viewOptions;
        console.log("this.views", this.views)
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);        
        this.render();       
    });*/
    if(this.properties.list){
      this.loadListColumns().then((results)=>{
        this.colInternalName = results;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);        
        this.render();
      });
    }
    //this.isListDataSet();
  } 
  /*protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'listName' && newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.close();
      this.context.propertyPane.open();
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, oldValue);
    }
  }*/
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    
    console.log("onPropertyPaneFieldChanged");
    console.log("this.colInternalName", this.colInternalName);
    console.log("this.properties", this.properties);
    
    if (this.properties.list){
      console.log("if (this.properties.list)");
      this.loadListColumns().then((results)=>{
        this.colInternalName = results;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);        
        this.render()
      });
    }

    // this.context.propertyPane.close();
    // this.context.propertyPane.open();
  }
  private loadViewsButtonClick(oldVal: any): any {   
    this.context.propertyPane.close();
    this.context.propertyPane.open();
  }  
  /* Loading Dpd with list names - End */

  private isManagedDataBtnDisabled: any;
  private isListDataSet(): void{
    if (this.properties.list && this.properties.siteUrl && this.properties.view) this.isManagedDataBtnDisabled = false;
    else this.isManagedDataBtnDisabled = true;
    this.context.propertyPane.refresh();
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);        
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Fill in the list properties to populate the data'
          },
          groups: [            
            {
              groupName: 'List Information',
              groupFields: [
                PropertyPaneTextField('siteUrl', {
                  label: 'Site URL',
                  value: this.properties.siteUrl,
                }),
                // PropertyPaneTextField('listName', {
                //   label: 'List Name',
                //   value: this.properties.listName
                // }),
                // PropertyPaneButton('loadViews',  {  
                //   text: "Load Views",  
                //   buttonType: PropertyPaneButtonType.Normal,  
                //   onClick: this.loadViewsButtonClick.bind(this)  
                // }),  
                // PropertyPaneDropdown('viewName', {
                //   label: 'View',
                //   options: this.views,
                //   selectedKey: this.properties.viewName
                // }),
                // PropertyFieldSitePicker('sites', {
                //   label: 'Select sites',
                //   initialSites: this.properties.sites,
                //   context: this.context as any,
                //   deferredValidationTime: 500,
                //   multiSelect: false,
                //   onPropertyChange: this.onPropertyPaneFieldChanged,
                //   properties: this.properties,
                //   key: 'sitesFieldId'
                // }),
                PropertyFieldListPicker('list', {
                  label: 'Select a list',
                  selectedList: this.properties.list,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId',
                  webAbsoluteUrl: this.properties.siteUrl
                }),
                PropertyFieldViewPicker('view', {
                  label: 'Select a view',
                  listId: this.properties.list,
                  selectedView: this.properties.view,
                  orderBy: PropertyFieldViewPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'viewPickerFieldId',
                  webAbsoluteUrl:this.properties.siteUrl
                }),
                PropertyPaneTextField('numItems', {
                  label: 'Number of Items',
                  value: this.properties.numItems,
                }),
                PropertyFieldCollectionData('collectionData', {
                  key: "collectionData",
                  label: "List View Collection data",
                  panelHeader: "List View Collection data",
                  manageBtnLabel: "Manage collection data",
                  value: this.properties.collectionData,
                  enableSorting: true,
                  panelDescription: "Please make sure to set up all the list data from the web part properties before setting in the columns to be displayed.",
                  fields: [
                    {id: "isStatus", title:"Status", type: CustomCollectionFieldType.boolean, required: false, defaultValue: false},
                    {id: "displayName", title:"Header Display Name", type: CustomCollectionFieldType.string, required: true},
                    {id: "fieldName", title:"Field (Internal Name)", type: CustomCollectionFieldType.dropdown, options: this.colInternalName, required: true},
                    {id: "isLink", title:"Hyperlink", type: CustomCollectionFieldType.boolean, required: false, defaultValue: false},
                    {id: "urlFieldName", title:"Link Field (Internal Name)", type: CustomCollectionFieldType.dropdown, options: this.colInternalName, required: false},
                    {id: "minWidth", title:"Min Width", type: CustomCollectionFieldType.number, required: true,},
                    {id: "maxWidth", title:"Max Width", type: CustomCollectionFieldType.number, required: true},
                    {id: "sorting", title:"Sorting", type: CustomCollectionFieldType.boolean, required: false, defaultValue: true},
                    {id: "isResizable", title:"Resizable", type: CustomCollectionFieldType.boolean, required: false, defaultValue: false},
                    {id: "isDate", title:"Date", type: CustomCollectionFieldType.boolean, required: false, defaultValue: false},
                    {id: "isTotalCost", title:"Total Cost", type: CustomCollectionFieldType.boolean, required: false, defaultValue: false},
                  ],
                  disabled: false
                }),
              ]
            },
            {
              groupName: 'Data Display',
              groupFields:[
                PropertyPaneTextField('instructionText', {
                  label: 'Instruction Text',
                  value: this.properties.instructionText,
                }),
                PropertyPaneTextField('filterPlaceholder', {
                  label: 'Search Placeholder Text',
                  value: this.properties.filterPlaceholder,
                }),
                PropertyPaneCheckbox('showFilter', {
                  text: 'Show Search/Filter',
                  checked: this.properties.showFilter
                }),
                PropertyPaneCheckbox('showSelectedItemsCount', {
                  text: 'Show Selected Items Count',
                  checked: this.properties.showSelectedItemsCount
                }),
                PropertyPaneCheckbox('showItemsCount', {
                  text: 'Show Items Count',
                  checked: this.properties.showItemsCount
                }),
                PropertyPaneCheckbox('showTotalCost', {
                  text: 'Show Total Cost',
                  checked: this.properties.showTotalCost
                }),
              ]
            },
          ]
        }
      ]
    };
  }
}
