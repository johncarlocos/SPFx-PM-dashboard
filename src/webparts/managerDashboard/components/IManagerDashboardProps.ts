import { SPHttpClient } from '@microsoft/sp-http';

export interface IManagerDashboardProps {
  siteUrl: string;
  userDisplayName: string;
  spHttpClient: SPHttpClient;
}
