import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
export interface IAdFortrexWebpartWebPartProps {
    description: string;
    /** Dataverse environment root URL, e.g. https://org1bec48f0.crm8.dynamics.com */
    dataverseUrl: string;
}
export default class AdFortrexWebpartWebPart extends BaseClientSideWebPart<IAdFortrexWebpartWebPartProps> {
    private _isDarkTheme;
    private _initError;
    private _serviceError;
    private _dataverseService;
    protected onInit(): Promise<void>;
    /** (Re-)creates the DataverseService for the currently configured URL. */
    private _initService;
    /** Re-init and re-render when the property pane value changes. */
    protected onPropertyPaneFieldChanged(propertyPath: string): Promise<void>;
    render(): void;
    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void;
    protected onDispose(): void;
    /** Render edge-to-edge on Team Sites */
    protected get isFullBleed(): boolean;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=AdFortrexWebpartWebPart.d.ts.map