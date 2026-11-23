import * as React from 'react';
import { ISupplier } from '../../models';
import { DataverseService, ODataRecord } from '../../services/DataverseService';
export interface INewRequestFormProps {
    dataverseService: DataverseService;
    context: import('@microsoft/sp-webpart-base').WebPartContext;
    /** Active role selected in AppShell header (e.g. requester, manager, approver_tier1/2/3) */
    activeRole?: string;
    /** Optional request record to load in edit / revise mode (e.g., Draft or Returned) */
    initialRecord?: ODataRecord | null;
    /** Callback when user finishes editing, submits, or cancels */
    onFinished?: () => void;
    /** Callback notifying parent when form has unsaved/dirty data */
    onDirtyChange?: (isDirty: boolean) => void;
}
interface IAttachment {
    name: string;
    mimeType: string;
    base64: string;
    sizeKb: number;
}
interface INewRequestFormState {
    sites: ODataRecord[];
    suppliers: ISupplier[];
    loadingMeta: boolean;
    metaError: string | null;
    sitesError: string | null;
    requesterGuid: string | null;
    requesterName: string;
    requesterEmail: string;
    userLookupError: string | null;
    managerGuid: string | null;
    managerName: string;
    managerEmail: string;
    managerNavProp: string;
    businessUnit: string;
    selectedSiteId: string;
    selectedSiteName: string;
    autoCostCentre: string;
    autoDivision: string;
    autoBudgetOwner: string;
    sctFunction: string;
    costIsFor: string;
    requestTitle: string;
    description: string;
    purchaseType: string;
    isNewSupplier: boolean;
    selectedSupplierId: string;
    newSupplierName: string;
    newSupplierLegalName: string;
    newSupplierContactName: string;
    newSupplierContactEmail: string;
    newSupplierContactPhone: string;
    newSupplierW9Notes: string;
    estimatedAmount: string;
    termStartDate: Date | null;
    termEndDate: Date | null;
    pricingBasis: string;
    dateNeededBy: Date | null;
    urgency: string;
    urgencyReason: string;
    emergencyDeclaredBy: string;
    deliverTo: string;
    receivingContact: string;
    equipmentAssetTag: string;
    costRechargedToCustomer: boolean;
    customerName: string;
    customerAgreementAttached: boolean;
    techOrDataInvolved: boolean;
    supplierOnPremises: boolean;
    possibleCapitalPurchase: boolean;
    isConfidential: boolean;
    isRetrospective: boolean;
    declarationNoPriorCommitment: boolean;
    declarationNoConflictOfInterest: boolean;
    attachments: IAttachment[];
    uploadingFiles: boolean;
    submitting: boolean;
    isSubmitting: boolean;
    savingDraft: boolean;
    submitError: string | null;
    submitSuccess: string | null;
    savedId: string | null;
    errors: Record<string, string>;
    currentStep: number;
}
export default class NewRequestForm extends React.Component<INewRequestFormProps, INewRequestFormState> {
    private _service;
    private _fileInputRef;
    constructor(props: INewRequestFormProps);
    componentDidMount(): void;
    componentWillUnmount(): void;
    componentDidUpdate(prevProps: INewRequestFormProps, prevState: INewRequestFormState): void;
    private _isDirty;
    private _handleBeforeUnload;
    /** Resolves workflow routing based on submitter's active authority / role */
    private _getRoutingConfig;
    private _loadMeta;
    /** Resolve site display name — PrimaryNameAttribute confirmed as sct_site_name */
    private _siteLabel;
    /** Resolve site primary key — PrimaryIdAttribute confirmed as sct_sites_masterid */
    private _siteId;
    private _onSiteSelected;
    private _validate;
    private _buildPayload;
    private _submit;
    private _onFilesSelected;
    private _removeAttachment;
    private _err;
    private _clrErr;
    private _renderHeaderSection;
    private _renderRequirementSection;
    private _renderFinancialSection;
    private _renderLogisticsSection;
    private _renderDeclarationsSection;
    private _renderAttachmentsSection;
    private _canGoNext;
    private _renderReviewStep;
    private _renderSuccess;
    render(): React.ReactElement;
}
export {};
//# sourceMappingURL=NewRequestForm.d.ts.map