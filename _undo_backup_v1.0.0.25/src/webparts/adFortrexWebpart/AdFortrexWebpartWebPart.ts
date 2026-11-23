import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { AadHttpClient } from '@microsoft/sp-http';

import AppShell from '../../components/shell/AppShell';
import { IAppShellProps } from '../../components/shell/IAppShellProps';
import { DataverseService } from '../../services/DataverseService';
import ErrorBoundary from '../../components/common/ErrorBoundary';

export interface IAdFortrexWebpartWebPartProps {
  description: string;
  /** Dataverse environment root URL, e.g. https://org1bec48f0.crm8.dynamics.com */
  dataverseUrl: string;
}

export default class AdFortrexWebpartWebPart extends BaseClientSideWebPart<IAdFortrexWebpartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _initError: string = '';
  private _serviceError: string = '';
  private _dataverseService: DataverseService | undefined = undefined;

  // ── Lifecycle ────────────────────────────────────────────────────────────────

  protected async onInit(): Promise<void> {
    await super.onInit();
    await this._initService();
  }

  /** (Re-)creates the DataverseService for the currently configured URL. */
  private async _initService(): Promise<void> {
    const url = (this.properties.dataverseUrl || '').trim();
    if (!url) {
      this._dataverseService = undefined;
      this._initError = '';
      return;
    }
    try {
      const aadClient: AadHttpClient = await this.context.aadHttpClientFactory
        .getClient(url);
      this._dataverseService = new DataverseService(aadClient, url);
      this._initError = '';
    } catch (err: unknown) {
      this._dataverseService = undefined;
      this._initError = err instanceof Error
        ? err.message
        : 'Failed to initialise Dataverse connection.';
      this._serviceError = '';
      return;
    }

    // Test whether the SCT tables actually exist in this environment
    try {
      await this._dataverseService.testConnection();
      this._serviceError = '';
    } catch (err: unknown) {
      this._serviceError = err instanceof Error
        ? err.message
        : 'Dataverse tables could not be reached in this environment.';
    }
  }

  /** Re-init and re-render when the property pane value changes. */
  protected async onPropertyPaneFieldChanged(propertyPath: string): Promise<void> {
    if (propertyPath === 'dataverseUrl') {
      await this._initService();
      this.render();
    }
  }

  public render(): void {
    const url = (this.properties.dataverseUrl || '').trim();

    // ── Not yet configured ──────────────────────────────────────────────────
    if (!url) {
      ReactDom.render(
        React.createElement('div', {
          style: {
            margin: 24, padding: 20, background: '#EFF6FC',
            border: '1px solid #0078D4', borderRadius: 8, fontFamily: 'Segoe UI, sans-serif',
          },
        },
          React.createElement('strong', { style: { color: '#004578', fontSize: 14 } },
            'Fortrex SCT — Configuration required'
          ),
          React.createElement('p', { style: { color: '#004578', marginTop: 8, fontSize: 13 } },
            'Please open the web part property pane (edit the page → click the web part → click the pencil icon) and enter your Dataverse Environment URL, e.g. https://orgxxxxxxxx.crm8.dynamics.com'
          )
        ),
        this.domElement
      );
      return;
    }

    // ── AAD / connection error ──────────────────────────────────────────────
    if (this._initError || !this._dataverseService) {
      ReactDom.render(
        React.createElement('div', {
          style: {
            margin: 24, padding: 24, background: '#FFF4CE',
            border: '1px solid #F7C948', borderRadius: 8, fontFamily: 'Segoe UI, sans-serif',
          },
        },
          React.createElement('strong', { style: { color: '#6B3A00' } },
            'Fortrex SCT — Initialisation Error'
          ),
          React.createElement('p', { style: { color: '#6B3A00', marginTop: 8, fontSize: 13 } },
            this._initError ||
            'Dataverse service unavailable. Ensure the AAD API permission for this URL has been approved in SharePoint Admin Centre → API Access.'
          )
        ),
        this.domElement
      );
      return;
    }

    const element = React.createElement(
      ErrorBoundary, {},
      React.createElement<IAppShellProps>(AppShell, {
        context:          this.context,
        dataverseService: this._dataverseService,
        userDisplayName:  this.context.pageContext.user.displayName,
        isDarkTheme:      this._isDarkTheme,
        description:      this.properties.description,
        serviceError:     this._serviceError || undefined,
      })
    );

    ReactDom.render(element, this.domElement);
  }

  // ── Theme & disposal ─────────────────────────────────────────────────────────

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link',     semanticColors.link     || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /** Render edge-to-edge on Team Sites */
  protected get isFullBleed(): boolean {
    return true;
  }

  // ── Property pane ────────────────────────────────────────────────────────────

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Fortrex SCT Web Part Settings' },
          groups: [
            {
              groupName: 'Dataverse Connection',
              groupFields: [
                PropertyPaneTextField('dataverseUrl', {
                  label: 'Dataverse Environment URL',
                  description: 'Root URL of your Power Platform environment, e.g. https://orgxxxxxxxx.crm8.dynamics.com',
                  placeholder: 'https://orgxxxxxxxx.crm8.dynamics.com',
                  onGetErrorMessage: (value: string) => {
                    if (!value || !value.trim()) return 'Dataverse URL is required.';
                    if (!/^https:\/\/.+\.dynamics\.com$/i.test(value.trim())) {
                      return 'Must be a valid Dynamics 365 URL (https://…dynamics.com).';
                    }
                    return '';
                  },
                  deferredValidationTime: 800,
                })
              ]
            },
            {
              groupName: 'General',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Description',
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
