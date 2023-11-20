import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMrfActionsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  listName: string;
  siteUrl: string;
  numItems: string;
  viewName: string;
  statusCol: string;

  columnsJSON: string;
  collectionData: any[];
}
