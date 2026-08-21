export interface ISubscriptionUpdatesProps {
  siteUrl: string;
  listId: string;
  subscribeForColumn: string;
  subscribedByColumn: string;
  subscribedOnColumn: string;
  referenceIdColumn?: string;
  customTitle: string;
  showTitle: boolean;
  descriptionText: string;
  successNotification: string;
  buttonText: string;
  showSuccessNotification: boolean;
  titleFontSize: number;
  descriptionFontSize: number;
  notificationFontSize: number;
  buttonFontSize: number;
  choicesFontSize: number;
  choicesLabel: string;
  titleBarOpacity: number;
  sp: any; // PnPjs instance for the selected site
  userDisplayName: string;
  userEmail: string;
}
