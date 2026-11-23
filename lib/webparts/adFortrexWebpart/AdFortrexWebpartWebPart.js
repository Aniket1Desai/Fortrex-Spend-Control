import { __awaiter, __extends, __generator } from "tslib";
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import AppShell from '../../components/shell/AppShell';
import { DataverseService } from '../../services/DataverseService';
import ErrorBoundary from '../../components/common/ErrorBoundary';
var AdFortrexWebpartWebPart = /** @class */ (function (_super) {
    __extends(AdFortrexWebpartWebPart, _super);
    function AdFortrexWebpartWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._initError = '';
        _this._serviceError = '';
        _this._dataverseService = undefined;
        return _this;
    }
    // ── Lifecycle ────────────────────────────────────────────────────────────────
    AdFortrexWebpartWebPart.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _super.prototype.onInit.call(this)];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, this._initService()];
                    case 2:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /** (Re-)creates the DataverseService for the currently configured URL. */
    AdFortrexWebpartWebPart.prototype._initService = function () {
        return __awaiter(this, void 0, void 0, function () {
            var url, aadClient, err_1, err_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        url = (this.properties.dataverseUrl || '').trim();
                        if (!url) {
                            this._dataverseService = undefined;
                            this._initError = '';
                            return [2 /*return*/];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this.context.aadHttpClientFactory
                                .getClient(url)];
                    case 2:
                        aadClient = _a.sent();
                        this._dataverseService = new DataverseService(aadClient, url);
                        this._initError = '';
                        return [3 /*break*/, 4];
                    case 3:
                        err_1 = _a.sent();
                        this._dataverseService = undefined;
                        this._initError = err_1 instanceof Error
                            ? err_1.message
                            : 'Failed to initialise Dataverse connection.';
                        this._serviceError = '';
                        return [2 /*return*/];
                    case 4:
                        _a.trys.push([4, 6, , 7]);
                        return [4 /*yield*/, this._dataverseService.testConnection()];
                    case 5:
                        _a.sent();
                        this._serviceError = '';
                        return [3 /*break*/, 7];
                    case 6:
                        err_2 = _a.sent();
                        this._serviceError = err_2 instanceof Error
                            ? err_2.message
                            : 'Dataverse tables could not be reached in this environment.';
                        return [3 /*break*/, 7];
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    /** Re-init and re-render when the property pane value changes. */
    AdFortrexWebpartWebPart.prototype.onPropertyPaneFieldChanged = function (propertyPath) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(propertyPath === 'dataverseUrl')) return [3 /*break*/, 2];
                        return [4 /*yield*/, this._initService()];
                    case 1:
                        _a.sent();
                        this.render();
                        _a.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        });
    };
    AdFortrexWebpartWebPart.prototype.render = function () {
        var url = (this.properties.dataverseUrl || '').trim();
        // ── Not yet configured ──────────────────────────────────────────────────
        if (!url) {
            ReactDom.render(React.createElement('div', {
                style: {
                    margin: 24, padding: 20, background: '#EFF6FC',
                    border: '1px solid #0078D4', borderRadius: 8, fontFamily: 'Segoe UI, sans-serif',
                },
            }, React.createElement('strong', { style: { color: '#004578', fontSize: 14 } }, 'Fortrex SCT — Configuration required'), React.createElement('p', { style: { color: '#004578', marginTop: 8, fontSize: 13 } }, 'Please open the web part property pane (edit the page → click the web part → click the pencil icon) and enter your Dataverse Environment URL, e.g. https://orgxxxxxxxx.crm8.dynamics.com')), this.domElement);
            return;
        }
        // ── AAD / connection error ──────────────────────────────────────────────
        if (this._initError || !this._dataverseService) {
            ReactDom.render(React.createElement('div', {
                style: {
                    margin: 24, padding: 24, background: '#FFF4CE',
                    border: '1px solid #F7C948', borderRadius: 8, fontFamily: 'Segoe UI, sans-serif',
                },
            }, React.createElement('strong', { style: { color: '#6B3A00' } }, 'Fortrex SCT — Initialisation Error'), React.createElement('p', { style: { color: '#6B3A00', marginTop: 8, fontSize: 13 } }, this._initError ||
                'Dataverse service unavailable. Ensure the AAD API permission for this URL has been approved in SharePoint Admin Centre → API Access.')), this.domElement);
            return;
        }
        var element = React.createElement(ErrorBoundary, {}, React.createElement(AppShell, {
            context: this.context,
            dataverseService: this._dataverseService,
            userDisplayName: this.context.pageContext.user.displayName,
            isDarkTheme: this._isDarkTheme,
            description: this.properties.description,
            serviceError: this._serviceError || undefined,
        }));
        ReactDom.render(element, this.domElement);
    };
    // ── Theme & disposal ─────────────────────────────────────────────────────────
    AdFortrexWebpartWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme)
            return;
        this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors;
        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }
    };
    AdFortrexWebpartWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(AdFortrexWebpartWebPart.prototype, "isFullBleed", {
        /** Render edge-to-edge on Team Sites */
        get: function () {
            return true;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(AdFortrexWebpartWebPart.prototype, "dataVersion", {
        // ── Property pane ────────────────────────────────────────────────────────────
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    AdFortrexWebpartWebPart.prototype.getPropertyPaneConfiguration = function () {
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
                                    onGetErrorMessage: function (value) {
                                        if (!value || !value.trim())
                                            return 'Dataverse URL is required.';
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
    };
    return AdFortrexWebpartWebPart;
}(BaseClientSideWebPart));
export default AdFortrexWebpartWebPart;
//# sourceMappingURL=AdFortrexWebpartWebPart.js.map