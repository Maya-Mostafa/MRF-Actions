import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneButtonType
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'MrfActionsWebPartStrings';
import MrfActions from './components/MrfActions';
import { IMrfActionsProps } from './components/IMrfActionsProps';

import { spfi, SPFx } from "@pnp/sp";
import { IField, IFieldInfo } from "@pnp/sp/fields/types";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/fields";

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';


export interface IMrfActionsWebPartProps {
  description: string;
  listName: string;
  numItems: string;
  viewName: string;
  siteUrl: string;
  collectionData: any[];
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
        listName: this.properties.listName,
        numItems: this.properties.numItems,
        viewName: this.properties.viewName,
        siteUrl: this.properties.siteUrl,
        collectionData: this.properties.collectionData
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
    try {
      const sp = spfi(this.properties.siteUrl).using(SPFx(this.context));  
      const listFields: any = await sp.web.lists.getByTitle(this.properties.listName).fields();

      console.log("all list fields", listFields);

      return listFields.filter((col: any)=>{
        // const excludedCols = ['_CheckinComment', '_ColorTag', '_CommentCount', '_ComplianceFlags', '_ComplianceTag', '_ComplianceTagUserId', '_ComplianceTagWrittenTime', '_CopySource', '_DisplayName', '_IsRecord', '_LikeCount', '_UIVersionString', 'Edit', 'Open','DocIcon','FileSizeDisplay','AppAuthor','AppEditor','ComplianceAssetId','ContentType','FolderChildCount','ItemChildCount','LinkFilenameNoMenu','MediaServiceImageTags','ParentLeafName','ParentVersionString'];
        const excludedCols = [''];
        //if (col.Hidden || excludedCols.indexOf(col.InternalName) !== -1) return false;
        return true;
      }).sort((a:any, b: any) => a.InternalName.localeCompare(b.InternalName)).map((col:any)=>{
        return {
          key: col.InternalName,
          text: col.InternalName
        }
      })

    }catch(error){
      console.log("error", error);
    }
  }

  /* Loading Dpd with list view names - Start */
  private views: IPropertyPaneDropdownOption[];
  private async loadListViews(): Promise<IPropertyPaneDropdownOption[]> {    
    console.log("loadListViews() --- ");
    const viewsTitle : any = [];
    try {
      const sp = spfi(this.properties.siteUrl).using(SPFx(this.context));  
      const listViews = await sp.web.lists.getByTitle(this.properties.listName).views();

      if (listViews) {
        listViews.map((result: any)=>{
          if (result.Title !== ''){
            viewsTitle.push({
              key: result.Title,
              text: result.Title
            });
          }
        });
        return viewsTitle;
      }
    } catch (error) {
      console.log("error", error);
    }
  }
  protected onPropertyPaneConfigurationStart(): void {
    console.log("onPropertyPaneConfigurationStart() --- ");
    // if (this.views) {
    //   this.render();  
    //   return;
    // }
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'views');
    this.loadListViews().then((viewOptions: IPropertyPaneDropdownOption[]): void => {
        this.views = viewOptions;
        console.log("this.views", this.views)
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);        
        this.render();       
    });
    this.loadListColumns().then((results)=>{
      this.colInternalName = results;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);        
      this.render();
    })

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
  private loadViewsButtonClick(oldVal: any): any {   
    this.context.propertyPane.close();
    this.context.propertyPane.open();
  }  
  /* Loading Dpd with list names - End */

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Select the list properties to populate the data'
          },
          groups: [
            {
              groupName: 'List Information',
              groupFields: [
                PropertyPaneTextField('siteUrl', {
                  label: 'Site URL',
                  value: this.properties.siteUrl,
                }),
                PropertyPaneTextField('listName', {
                  label: 'List Name',
                  value: this.properties.listName
                }),
                PropertyPaneButton('loadViews',  {  
                  text: "Load Views",  
                  buttonType: PropertyPaneButtonType.Normal,  
                  onClick: this.loadViewsButtonClick.bind(this)  
                }),  
                PropertyPaneDropdown('viewName', {
                  label: 'View',
                  options: this.views,
                  selectedKey: this.properties.viewName
                }),
                PropertyPaneTextField('numItems', {
                  label: 'Number of Items',
                  value: this.properties.numItems
                }),
                PropertyFieldCollectionData('collectionData', {
                  key: "collectionData",
                  label: "List View Collection data",
                  panelHeader: "List View Collection data - Please add the columns required to be displayed",
                  manageBtnLabel: "Manage collection data",
                  value: this.properties.collectionData,
                  enableSorting: true,
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
