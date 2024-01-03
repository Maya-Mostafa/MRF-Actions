import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMrfActionsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  siteUrl: string;
  numItems: string;
  collectionData: any[];

  list: string ;
  view: string;

  instructionText: string;
  showFilter: boolean;
  filterPlaceholder: string;
  showSelectedItemsCount: boolean;
  showItemsCount: boolean;
  showTotalCost: boolean;

  showRefresh: boolean;
  refreshText: string;
  refreshEvery5min: boolean;
}
