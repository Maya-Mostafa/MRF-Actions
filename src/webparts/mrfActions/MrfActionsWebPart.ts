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
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/views";

//import { FieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-controls-react/lib/FieldCollectionData';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';


export interface IMrfActionsWebPartProps {
  description: string;
  listName: string;
  numItems: string;
  viewName: string;
  siteUrl: string;
  statusCol: string;
  columnsJSON: string;
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
        statusCol: this.properties.statusCol,

        columnsJSON: this.properties.columnsJSON,
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
    this.loadListViews()
      .then((viewOptions: IPropertyPaneDropdownOption[]): void => {
        this.views = viewOptions;
        console.log("this.views", this.views)
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);        
        this.render();       
      });
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
  private ButtonClick(oldVal: any): any {   
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
                  onClick: this.ButtonClick.bind(this)  
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
                PropertyPaneTextField('statusCol', {
                  label: 'Status Column',
                  value: this.properties.statusCol,                  
                }),

                PropertyPaneTextField('columnsJSON', {
                  label: 'Columns JSON',
                  value: this.properties.columnsJSON,
                  multiline: true,
                  rows: 20
                }),
                PropertyFieldCollectionData('collectionData', {
                  key: "collectionData",
                  label: "Collection data",
                  panelHeader: "Collection data panel header",
                  manageBtnLabel: "Manage collection data",
                  value: this.properties.collectionData,
                  fields: [
                    {id: "fieldName", title:"Field Name", type: CustomCollectionFieldType.string, required: true},
                    {id: "displayName", title:"Display Name", type: CustomCollectionFieldType.string, required: true},
                    {id: "minWidth", title:"Min Column Width", type: CustomCollectionFieldType.number, required: true},
                    {id: "maxWidth", title:"Max Column Width", type: CustomCollectionFieldType.number, required: true},
                    {id: "sorting", title:"Sorting Column", type: CustomCollectionFieldType.boolean, required: false, defaultValue: true},
                    {id: "isResizable", title:"Resizable Column", type: CustomCollectionFieldType.boolean, required: false, defaultValue: false},
                    {id: "isDate", title:"Date field", type: CustomCollectionFieldType.boolean, required: false, defaultValue: false},
                    {id: "isTotalCost", title:"Total Cost field", type: CustomCollectionFieldType.boolean, required: false, defaultValue: false},
                    {id: "isStatus", title:"Status field", type: CustomCollectionFieldType.boolean, required: false, defaultValue: false}
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
