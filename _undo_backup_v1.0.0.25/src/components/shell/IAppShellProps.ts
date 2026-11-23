import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DataverseService } from '../../services/DataverseService';

export interface IAppShellProps {
  context: WebPartContext;
  dataverseService: DataverseService;
  userDisplayName: string;
  isDarkTheme: boolean;
  description: string;
  /** Set when Dataverse tables are unreachable — renders UI in disabled state with this message */
  serviceError?: string;
}
