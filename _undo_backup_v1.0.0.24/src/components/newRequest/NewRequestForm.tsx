import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { DayOfWeek } from '@fluentui/react/lib/Calendar';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Checkbox } from '@fluentui/react/lib/Checkbox';

import { ISupplier } from '../../models';
import { DataverseService, ODataRecord, strVal } from '../../services/DataverseService';
import { LoadingSpinner } from '../common';
import styles from '../shell/AppShell.module.scss';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import {
  getManager,
  sendManagerApprovalEmail,
  sendTierNotificationEmail,
  sendGraphEmail,
  generateApprovalEmailHTML,
} from '../../services/GraphMailService';

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// ── Prop / State types ──────────────────────────────────────────────────────────

export interface INewRequestFormProps {
  dataverseService: DataverseService;
  context: import('@microsoft/sp-webpart-base').WebPartContext;
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
  // Reference data — sites as raw OData records so field names are resolved at runtime
  sites: ODataRecord[];
  suppliers: ISupplier[];
  loadingMeta: boolean;
  metaError: string | null;
  sitesError: string | null;

  // Current user resolved from Dataverse systemuser table (for sct_Requester@odata.bind)
  requesterGuid: string | null;
  requesterName: string;
  requesterEmail: string;
  userLookupError: string | null;
  // Manager resolved from Graph API /me/manager → Dataverse systemusers
  managerGuid: string | null;
  managerName: string;
  managerEmail: string;
  // Actual OData navigation property name for the manager lookup — resolved from Dataverse metadata
  managerNavProp: string;

  // ── Section 1: Header / Requester
  businessUnit: string;
  selectedSiteId: string;
  selectedSiteName: string;
  autoCostCentre: string;
  autoDivision: string;
  autoBudgetOwner: string;
  sctFunction: string;
  costIsFor: string;

  // ── Section 2: Requirement & Supplier
  requestTitle: string;
  description: string;
  purchaseType: string;
  isNewSupplier: boolean;
  selectedSupplierId: string;
  // New supplier fields
  newSupplierName: string;
  newSupplierLegalName: string;
  newSupplierContactName: string;
  newSupplierContactEmail: string;
  newSupplierContactPhone: string;
  newSupplierW9Notes: string;

  // ── Section 3: Financials
  estimatedAmount: string;
  termStartDate: Date | null;
  termEndDate: Date | null;
  pricingBasis: string;
  dateNeededBy: Date | null;
  urgency: string;
  urgencyReason: string;
  emergencyDeclaredBy: string;

  // ── Section 4: Logistics
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

  // ── Section 5: Declarations
  declarationNoPriorCommitment: boolean;
  declarationNoConflictOfInterest: boolean;

  // ── Attachments
  attachments: IAttachment[];
  uploadingFiles: boolean;

  // ── Submit state
  submitting: boolean;
  isSubmitting: boolean;
  savingDraft: boolean;
  submitError: string | null;
  submitSuccess: string | null;
  savedId: string | null;

  // ── Validation errors
  errors: Record<string, string>;

  // ── Wizard step (0=General Info, 1=Financials, 2=Supplier & Compliance, 3=Review & Submit)
  currentStep: number;
}

// ── Choice option sets ──────────────────────────────────────────────────────────

const BUSINESS_UNIT_OPTIONS: IDropdownOption[] = [
  { key: 'PSSI', text: 'PSSI' },
  { key: 'PCI', text: 'PCI' },
  { key: 'SAFE', text: 'SAFE' },
  { key: 'CCCS', text: 'CCCS' },
];

const COST_IS_FOR_OPTIONS: IDropdownOption[] = [
  { key: 'This Site', text: 'This Site' },
  { key: 'A Customer Job', text: 'A Customer Job' },
  { key: 'Several Sites', text: 'Several Sites' },
  { key: 'Company-Wide', text: 'Company-Wide' },
];

const PURCHASE_TYPE_OPTIONS: IDropdownOption[] = [
  { key: 'Goods', text: 'Goods' },
  { key: 'Repair', text: 'Repair' },
  { key: 'Rental', text: 'Rental' },
  { key: 'Service', text: 'Service' },
  { key: 'Software', text: 'Software' },
  { key: 'Temporary Labor', text: 'Temporary Labor' },
  { key: 'Other', text: 'Other' },
];

const PRICING_BASIS_OPTIONS: IDropdownOption[] = [
  { key: 'Fixed Price', text: 'Fixed Price' },
  { key: 'Rate Card', text: 'Rate Card' },
  { key: 'Time and Materials with Cap', text: 'Time and Materials with Cap' },
  { key: 'Usage-Based Estimate', text: 'Usage-Based Estimate' },
];

const URGENCY_OPTIONS: IDropdownOption[] = [
  { key: 'Standard', text: 'Standard' },
  { key: 'Urgent', text: 'Urgent' },
  { key: 'Emergency', text: 'Emergency' },
];

// ── Helpers ────────────────────────────────────────────────────────────────────

// ── Dataverse choice-field integer mappings ────────────────────────────────────
// Choice fields in Dataverse are Edm.Int32. Publisher prefix sct_ uses base 854000000.
// Confirmed from actual Dataverse record: sct_urgency=854000000→"Standard",
// sct_purchase_type=854000000→"Goods", sct_business_unit=854000000→"PSSI", etc.
const URGENCY_INT: Record<string, number> = {
  'Standard': 854000000,
  'Urgent': 854000001,
  'Emergency': 854000002,
};
const PURCHASE_TYPE_INT: Record<string, number> = {
  'Goods': 854000000,
  'Repair': 854000001,
  'Rental': 854000002,
  'Service': 854000003,
  'Software': 854000004,
  'Temporary Labor': 854000005,
  'Other': 854000006,
};
const PRICING_BASIS_INT: Record<string, number> = {
  'Fixed Price': 854000000,
  'Rate Card': 854000001,
  'Time and Materials with Cap': 854000002,
  'Usage-Based Estimate': 854000003,
};
const COST_IS_FOR_INT: Record<string, number> = {
  'This Site': 854000000,
  'A Customer Job': 854000001,
  'Several Sites': 854000002,
  'Company-Wide': 854000003,
};
const BUSINESS_UNIT_INT: Record<string, number> = {
  'PSSI': 854000000,
  'PCI': 854000001,
  'SAFE': 854000002,
  'CCCS': 854000003,
};

/** Returns the integer option-set value if known; otherwise undefined (field not sent). */
function choiceInt(map: Record<string, number>, label: string): number | undefined {
  return map[label];
}

function readFileAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      resolve(result.split(',')[1] ?? result);
    };
    reader.onerror = () => reject(new Error(`Failed to read ${file.name}`));
    reader.readAsDataURL(file);
  });
}


/** Edm.Date — Dataverse DateOnly columns require strict YYYY-MM-DD, no time component.
 *  Uses local date components to avoid UTC midnight shift (e.g. Oct 31 → Oct 30).
 *  Avoids padStart — not available in the project's TS lib target. */
function dateOnly(d: Date | null): string | undefined {
  if (!d) return undefined;
  const y = d.getFullYear();
  const mo = d.getMonth() + 1;
  const da = d.getDate();
  const mm = mo < 10 ? ('0' + mo) : String(mo);
  const dd = da < 10 ? ('0' + da) : String(da);
  return y + '-' + mm + '-' + dd;
}

// ── Section card wrapper ───────────────────────────────────────────────────────

interface ISectionProps {
  iconName: string;
  iconBg: string;
  iconColor: string;
  title: string;
  children: React.ReactNode;
  badge?: string;
}

const Section: React.FC<ISectionProps> = ({ iconName, iconBg, iconColor, title, badge, children }) => (
  <div className={styles.card} style={{ marginBottom: 16 }}>
    <div className={styles.cardHeader}>
      <div className={styles.cardTitle}>
        <div className={styles.cardTitleIcon} style={{ background: iconBg }} aria-hidden="true">
          <Icon iconName={iconName} style={{ color: iconColor, fontSize: 16 }} />
        </div>
        {title}
        {badge && (
          <span style={{ marginLeft: 10, background: iconBg, color: iconColor, fontSize: 11, fontWeight: 700, padding: '2px 10px', borderRadius: 12, letterSpacing: '0.3px' }}>
            {badge}
          </span>
        )}
      </div>
    </div>
    <div className={styles.cardBody}>{children}</div>
  </div>
);

const fieldStyles = { root: { marginBottom: 16 }, fieldGroup: { borderRadius: 6 } };
const dropdownStyles = { root: { marginBottom: 16 }, dropdown: { borderRadius: 6 } };
const checkStyles = { root: { marginBottom: 14 } };

// ── Row layout helper ─────────────────────────────────────────────────────────

const Row: React.FC<{ children: React.ReactNode; cols?: number }> = ({ children, cols = 2 }) => (
  <div style={{ display: 'grid', gridTemplateColumns: `repeat(${cols}, 1fr)`, gap: '0 24px' }}>
    {children}
  </div>
);

// ── Pre-approval guidance calculator ─────────────────────────────────────────

interface IApprovalGuidance {
  levels: string[];
  color: string;
  bg: string;
  iconName: string;
}

function calcApprovalGuidance(isNewSupplier: boolean, amount: number): IApprovalGuidance {
  if (isNewSupplier) {
    return {
      levels: ['L1 — Budget Owner', 'L2 — Senior Manager', 'L3 — ELT Leader'],
      color: '#A4262C', bg: '#FDE7E9', iconName: 'Warning',
    };
  }
  if (amount <= 2500) {
    return {
      levels: ['L1 — Budget Owner'],
      color: '#107C10', bg: '#DFF6DD', iconName: 'CheckMark',
    };
  }
  if (amount <= 10000) {
    return {
      levels: ['L2 — Senior Manager'],
      color: '#905400', bg: '#FFF4CE', iconName: 'Info',
    };
  }
  return {
    levels: ['L3 — ELT Leader'],
    color: '#A4262C', bg: '#FDE7E9', iconName: 'Warning',
  };
}

// ── Main component ─────────────────────────────────────────────────────────────

export default class NewRequestForm extends React.Component<INewRequestFormProps, INewRequestFormState> {
  private _service: DataverseService | null = null;
  private _fileInputRef: React.RefObject<HTMLInputElement> = React.createRef();

  constructor(props: INewRequestFormProps) {
    super(props);
    this.state = {
      sites: [], suppliers: [],
      loadingMeta: false, metaError: null, sitesError: null,
      requesterGuid: null, requesterName: '', requesterEmail: '', userLookupError: null,
      managerGuid: null, managerName: '', managerEmail: '', managerNavProp: 'cr2a4_Manager',

      businessUnit: '', selectedSiteId: '', selectedSiteName: '',
      autoCostCentre: '', autoDivision: '', autoBudgetOwner: '',
      sctFunction: '', costIsFor: '',

      requestTitle: '', description: '', purchaseType: '',
      isNewSupplier: false, selectedSupplierId: '',
      newSupplierName: '', newSupplierLegalName: '',
      newSupplierContactName: '', newSupplierContactEmail: '',
      newSupplierContactPhone: '', newSupplierW9Notes: '',

      estimatedAmount: '', termStartDate: null, termEndDate: null,
      pricingBasis: '', dateNeededBy: null,
      urgency: '', urgencyReason: '', emergencyDeclaredBy: '',

      deliverTo: '', receivingContact: '', equipmentAssetTag: '',
      costRechargedToCustomer: false, customerName: '',
      customerAgreementAttached: false, techOrDataInvolved: false,
      supplierOnPremises: false, possibleCapitalPurchase: false,
      isConfidential: false, isRetrospective: false,

      declarationNoPriorCommitment: false,
      declarationNoConflictOfInterest: false,

      attachments: [], uploadingFiles: false,
      submitting: false, isSubmitting: false, savingDraft: false,
      submitError: null, submitSuccess: null, savedId: null,
      errors: {},
      currentStep: 0,
    };
  }

  public componentDidMount(): void {
    this._service = this.props.dataverseService;
    void this._loadMeta();
    if (typeof window !== 'undefined') {
      window.addEventListener('beforeunload', this._handleBeforeUnload);
    }
  }

  public componentWillUnmount(): void {
    if (typeof window !== 'undefined') {
      window.removeEventListener('beforeunload', this._handleBeforeUnload);
    }
    if (this.props.onDirtyChange) {
      this.props.onDirtyChange(false);
    }
  }

  public componentDidUpdate(prevProps: INewRequestFormProps, prevState: INewRequestFormState): void {
    if (this.props.onDirtyChange) {
      const wasDirty = this._isDirty(prevState);
      const isNowDirty = this._isDirty(this.state);
      if (wasDirty !== isNowDirty) {
        this.props.onDirtyChange(isNowDirty);
      }
    }
  }

  private _isDirty(s: INewRequestFormState): boolean {
    if (s.submitSuccess || s.savedId) return false;
    if (s.currentStep > 0) return true;
    if ((s.requestTitle || '').trim().length > 0) return true;
    if ((s.description || '').trim().length > 0) return true;
    if ((s.purchaseType || '').length > 0) return true;
    if ((s.businessUnit || '').length > 0) return true;
    if ((s.selectedSiteId || '').length > 0) return true;
    if ((s.estimatedAmount || '').trim().length > 0) return true;
    if ((s.urgency || '').length > 0) return true;
    if ((s.newSupplierName || '').trim().length > 0) return true;
    if ((s.selectedSupplierId || '').length > 0) return true;
    if ((s.attachments || []).length > 0) return true;
    if ((s.deliverTo || '').trim().length > 0) return true;
    if ((s.receivingContact || '').trim().length > 0) return true;
    if ((s.customerName || '').trim().length > 0) return true;
    if ((s.equipmentAssetTag || '').trim().length > 0) return true;
    if ((s.sctFunction || '').trim().length > 0) return true;
    if ((s.costIsFor || '').length > 0) return true;
    return false;
  }

  private _handleBeforeUnload = (e: BeforeUnloadEvent): void => {
    if (this._isDirty(this.state)) {
      e.preventDefault();
      e.returnValue = '';
    }
  };

  // ── Data loading ─────────────────────────────────────────────────────────────

  private async _loadMeta(): Promise<void> {
    if (!this._service) return;
    this.setState({ loadingMeta: true, metaError: null, sitesError: null, userLookupError: null });
    try {
      // Resolve current user as a Dataverse systemuser for sct_Requester@odata.bind.
      // Primary: use AAD Object ID from SPFx context (exact match).
      // Fallback: look up by email if AAD ID is unavailable.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ctx = this.props.context as any;
      const aadId: string | undefined = ctx?.pageContext?.aadInfo?.userId?.toString();
      const spEmail: string | undefined = ctx?.pageContext?.user?.email;

      const resolveUser = async (): Promise<{ guid: string | null; name: string; email: string; err: string | null }> => {
        if (aadId) {
          const u = await this._service!.getCurrentSystemUser(aadId);
          if (u) return { guid: u.systemuserid, name: u.fullname, email: u.internalemailaddress, err: null };
        }
        if (spEmail) {
          const u = await this._service!.getSystemUserByEmail(spEmail);
          if (u) return { guid: u.systemuserid, name: u.fullname, email: u.internalemailaddress, err: null };
        }
        return {
          guid: null,
          name: ctx?.pageContext?.user?.displayName ?? '',
          email: spEmail ?? '',
          err: 'Could not resolve your Dataverse user record — requester will not be bound.',
        };
      };

      const [sites, suppliers, userResult, managerNavProp] = await Promise.all([
        this._service.getSitesRaw(true).catch((e: unknown): ODataRecord[] => {
          const msg = e instanceof Error ? e.message : 'Failed to load sites';
          this.setState({ sitesError: msg });
          return [];
        }),
        this._service.getSuppliers(true).catch((): ISupplier[] => []),
        resolveUser(),
        // cr2a4_manager = logical name of the Manager lookup (publisher prefix cr2a4_).
        // getNavPropName returns: (1) relationship nav prop, (2) attribute SchemaName 'cr2a4_Manager',
        // (3) the logical name as last resort. The .catch() hardcodes the known schema name.
        this._service.getNavPropName('sct_request', 'cr2a4_manager').catch(() => 'cr2a4_Manager'),
      ]);

      // ── Resolve manager GUID ───────────────────────────────────────────────────
      //
      // Resolution order (stops at first successful match):
      //   1. Graph API /me/manager → manager's AAD Object ID → Dataverse systemuserid
      //   2. Graph API /me/manager → manager's email → Dataverse systemuserid
      //   3. Dataverse systemuser.parentsystemuserid (org chart sync fallback)
      //
      // ── Resolve manager via Graph API /me/manager and Dataverse systemuser ──
      let managerGuid: string | null = null;
      let managerName = '';
      let managerEmail = '';

      try {
        const graphClient: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient('3');
        // Phase 2 Step 1: Call await graphClient.api('/me/manager').version('v1.0').get()
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const managerData: any = await graphClient.api('/me/manager').version('v1.0').get().catch(() => null);
        if (managerData) {
          // Extract manager email (check mail or userPrincipalName)
          managerEmail = (managerData.mail || managerData.userPrincipalName || '').trim();
          managerName = managerData.displayName || managerEmail;

          // Phase 2 Step 2: Query Dataverse systemusers table
          // GET /api/data/v9.2/systemusers?$select=systemuserid,fullname&$filter=internalemailaddress eq '${managerEmail}'
          if (managerEmail && this._service) {
            managerGuid = await this._service.getSystemUserIdByEmail(managerEmail).catch(() => null);
          }
          // Secondary fallback by AAD Object ID if email had no direct match in Dataverse
          if (!managerGuid && managerData.id && this._service) {
            const mgr = await this._service.getCurrentSystemUser(managerData.id).catch(() => null);
            managerGuid = mgr?.systemuserid ?? null;
          }
          // Tertiary fallback: match by email prefix or display name in systemusers
          if (!managerGuid && managerEmail && this._service) {
            const mgr = await this._service.getSystemUserByEmail(managerEmail).catch(() => null);
            managerGuid = mgr?.systemuserid ?? null;
          }
        }
      } catch (graphErr: unknown) {
        console.warn('[NewRequestForm] Graph /me/manager lookup warning:', graphErr);
      }

      // Fallback: Dataverse parentsystemuserid (works when Graph is unavailable or in dev/test)
      if (!managerGuid && userResult.guid && this._service) {
        const u = await this._service.getCurrentSystemUser(aadId ?? '').catch(() => null) ??
          (spEmail ? await this._service.getSystemUserByEmail(spEmail).catch(() => null) : null);
        managerGuid = u?._parentsystemuserid_value ?? null;
        if (!managerGuid) {
          managerGuid = await this._service.getManagerSystemUserId(userResult.guid).catch(() => null);
        }
        if (managerGuid) {
          const sysMgr = await this._service.getSystemUserById(managerGuid).catch(() => null);
          if (sysMgr) {
            if (!managerName) managerName = sysMgr.fullname;
            if (!managerEmail) managerEmail = sysMgr.internalemailaddress;
          }
        }
      }

      let initialUpdate: Partial<INewRequestFormState> = {};
      const rec = this.props.initialRecord;
      if (rec) {
        const id = String(rec['sct_requestid'] ?? '');
        const siteId = strVal(rec, '_sct_site_value') || strVal(rec, 'sct_site');
        const foundSite = sites.find(s => this._siteId(s) === siteId);
        initialUpdate = {
          savedId: id || null,
          requestTitle: strVal(rec, 'sct_request_title') || strVal(rec, 'sct_title'),
          description: strVal(rec, 'sct_description'),
          businessUnit: strVal(rec, 'sct_business_unit'),
          selectedSiteId: siteId,
          selectedSiteName: foundSite ? this._siteLabel(foundSite) : '',
          autoCostCentre: strVal(rec, 'sct_cost_center') || (foundSite ? strVal(foundSite, 'sct_costcenter') : ''),
          autoDivision: strVal(rec, 'sct_division') || (foundSite ? strVal(foundSite, 'sct_region') : ''),
          autoBudgetOwner: strVal(rec, 'sct_budgetowner') || (foundSite ? strVal(foundSite, '_sct_budgetownerid_value') : ''),
          purchaseType: strVal(rec, 'sct_purchase_type'),
          isNewSupplier: Boolean(rec['sct_is_new_supplier']),
          selectedSupplierId: strVal(rec, '_sct_suggested_supplier_value'),
          newSupplierName: strVal(rec, 'sct_proposed_supplier_name'),
          estimatedAmount: strVal(rec, 'sct_estimated_total_value'),
          urgency: strVal(rec, 'sct_urgency'),
          urgencyReason: strVal(rec, 'sct_urgency_reason'),
          equipmentAssetTag: strVal(rec, 'sct_equipment_asset_tag'),
          costRechargedToCustomer: Boolean(rec['sct_cost_recharged_to_customer']),
          customerName: strVal(rec, 'sct_customer_name'),
          declarationNoPriorCommitment: Boolean(rec['sct_declaration_no_prior_commitment']),
          declarationNoConflictOfInterest: Boolean(rec['sct_declaration_no_conflict_of_interest']),
        };
      }

      this.setState({
        sites, suppliers, loadingMeta: false,
        requesterGuid: userResult.guid,
        requesterName: userResult.name,
        requesterEmail: userResult.email,
        userLookupError: userResult.err,
        managerGuid,
        managerName,
        managerEmail,
        managerNavProp: (managerNavProp && managerNavProp.toLowerCase() !== 'sct_manager') ? managerNavProp : 'cr2a4_Manager',
      }, () => {
        if (rec) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          this.setState(initialUpdate as any);
        }
      });
    } catch (err: unknown) {
      this.setState({
        metaError: err instanceof Error ? err.message : 'Failed to load form data',
        loadingMeta: false,
      });
    }
  }

  // ── Site selection — auto-fill cost centre / division / budget owner ──────────

  /** Resolve site display name — PrimaryNameAttribute confirmed as sct_site_name */
  private _siteLabel(site: ODataRecord): string {
    return strVal(site, 'sct_site_name') ||
      strVal(site, 'sct_name') ||
      strVal(site, 'sct_sitename') ||
      strVal(site, 'sct_sitecode') ||
      String(site['sct_sites_masterid'] ?? '');
  }

  /** Resolve site primary key — PrimaryIdAttribute confirmed as sct_sites_masterid */
  private _siteId(site: ODataRecord): string {
    return String(site['sct_sites_masterid'] ?? site['sct_sitesmasterid'] ?? '');
  }

  private _onSiteSelected(siteId: string): void {
    const site = this.state.sites.find(s => this._siteId(s) === siteId);
    if (!site) return;
    const rawBudgetOwner = strVal(site, '_sct_budgetownerid_value') ||
      strVal(site, 'sct_budgetowner') ||
      strVal(site, '_sct_budgetowner_value');
    this.setState({
      selectedSiteId: siteId,
      selectedSiteName: this._siteLabel(site),
      autoCostCentre: strVal(site, 'sct_costcenter') || strVal(site, 'sct_cost_center'),
      autoDivision: strVal(site, 'sct_region') || strVal(site, 'sct_division'),
      autoBudgetOwner: rawBudgetOwner,
      businessUnit: this.state.businessUnit || strVal(site, 'sct_businessunit'),
    });

    // If budget owner is a raw GUID, resolve display name from Dataverse systemusers
    const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (rawBudgetOwner && GUID_RE.test(rawBudgetOwner) && this._service) {
      void this._service.getSystemUserById(rawBudgetOwner).then(user => {
        if (user && user.fullname) {
          this.setState({ autoBudgetOwner: user.fullname });
        }
      }).catch(() => { /* non-fatal */ });
    }
  }

  // ── Validation ────────────────────────────────────────────────────────────────

  private _validate(): boolean {
    const s = this.state;
    const errs: Record<string, string> = {};

    // ── Section 1: Header ─────────────────────────────────────────────────────
    if (!s.businessUnit)
      errs['businessUnit'] = 'Business unit is required.';
    if (!s.selectedSiteId)
      errs['selectedSiteId'] = 'Site / Location is required.';

    // ── Section 2: Requirement ────────────────────────────────────────────────
    if (!s.description.trim())
      errs['description'] = 'Description / Justification is required.';
    if (!s.purchaseType)
      errs['purchaseType'] = 'Purchase type is required.';
    if (s.isNewSupplier && !s.newSupplierName.trim())
      errs['newSupplierName'] = 'Proposed supplier name is required for new supplier requests.';

    // ── Section 3: Financials ─────────────────────────────────────────────────
    const rawAmt = s.estimatedAmount ? String(s.estimatedAmount).trim() : '';
    const cleanAmt = rawAmt.replace(/[$,]/g, '').trim();
    const isValidPositiveNumber = /^[0-9]+(\.[0-9]{1,2})?$/.test(cleanAmt);
    const amt = parseFloat(cleanAmt);
    if (!rawAmt || isNaN(amt) || amt <= 0 || !isValidPositiveNumber)
      errs['estimatedAmount'] = 'Amount must be greater than zero and contain only valid positive numbers.';
    if (!s.dateNeededBy) {
      errs['dateNeededBy'] = 'Date needed by is required.';
    } else {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const neededDate = new Date(s.dateNeededBy);
      neededDate.setHours(0, 0, 0, 0);
      if (neededDate < today && !s.isRetrospective) {
        errs['dateNeededBy'] = 'Date needed by cannot be in the past (unless Retrospective Spend is declared).';
      }
    }
    if (!s.urgency)
      errs['urgency'] = 'Urgency is required.';
    if ((s.urgency === 'Urgent' || s.urgency === 'Emergency') && !s.urgencyReason.trim())
      errs['urgencyReason'] = 'Urgency reason is required for Urgent / Emergency requests.';
    if (s.urgency === 'Emergency' && !s.emergencyDeclaredBy.trim())
      errs['emergencyDeclaredBy'] = 'The person declaring the emergency must be named.';

    // ── Section 4: Logistics ──────────────────────────────────────────────────
    if ((s.purchaseType === 'Repair' || s.purchaseType === 'Rental') && !s.equipmentAssetTag.trim())
      errs['equipmentAssetTag'] = 'Asset tag / serial number is required for Repair or Rental.';
    if (s.costRechargedToCustomer && !s.customerName.trim())
      errs['customerName'] = 'Customer name is required when cost is recharged.';
    if (s.costRechargedToCustomer && !s.customerAgreementAttached)
      errs['customerAgreementAttached'] = 'Customer agreement or quote must be attached when cost is recharged to a customer.';

    // ── Section 5: Declarations (R1 + R10) ───────────────────────────────────
    if (!s.declarationNoPriorCommitment)
      errs['declarationNoPriorCommitment'] = 'You must confirm no prior commitment has been made (Rule R1).';
    if (!s.declarationNoConflictOfInterest)
      errs['declarationNoConflictOfInterest'] = 'You must declare no conflict of interest (Rule R10).';

    // ── Section 6: Approval evidence attachment ───────────────────────────────
    const validAmt = !isNaN(amt) && amt > 0 && isValidPositiveNumber;
    if ((validAmt || s.isNewSupplier) && s.attachments.length === 0) {
      const guidance = calcApprovalGuidance(s.isNewSupplier, validAmt ? amt : 0);
      errs['attachments'] = s.isNewSupplier
        ? 'Pre-approval evidence from L1, L2 and L3 approvers must be attached for new supplier requests.'
        : 'Written pre-approval evidence from ' + guidance.levels.join(' + ') + ' must be attached before submission.';
    }

    if (Object.keys(errs).length > 0) {
      const firstErrKey = errs['estimatedAmount'] ? 'estimatedAmount' : Object.keys(errs)[0];
      const firstErrMsg = errs[firstErrKey];
      this.setState({ errors: errs, submitError: `Validation error: ${firstErrMsg}` });
      return false;
    }

    this.setState({ errors: errs, submitError: null });
    return true;
  }

  // ── Payload builder ───────────────────────────────────────────────────────────

  private _buildPayload(isDraft: boolean): Record<string, unknown> {
    const s = this.state;
    const payload: Record<string, unknown> = {};

    // ── Text fields — names confirmed from actual Dataverse record ────────────
    // NOTE: sct_request has NO sct_name field. PrimaryNameAttribute is sct_sct_reference
    // (autonumber, auto-generated by Dataverse — never include it in the payload).
    // Build description — append new-supplier details inline (no dedicated schema fields for them)
    const descParts: string[] = [];
    if (s.description.trim()) descParts.push(s.description.trim());
    if (s.isNewSupplier && s.newSupplierName.trim()) {
      descParts.push(
        '[New Supplier] ' + s.newSupplierName.trim() +
        (s.newSupplierLegalName.trim() ? ' | Legal: ' + s.newSupplierLegalName.trim() : '') +
        (s.newSupplierContactName.trim() ? ' | Contact: ' + s.newSupplierContactName.trim() : '') +
        (s.newSupplierContactEmail.trim() ? ' <' + s.newSupplierContactEmail.trim() + '>' : '') +
        (s.newSupplierContactPhone.trim() ? ' ' + s.newSupplierContactPhone.trim() : '') +
        (s.newSupplierW9Notes.trim() ? ' | W-9: ' + s.newSupplierW9Notes.trim() : '')
      );
    }
    if (descParts.length) payload['sct_description'] = descParts.join('\n\n');
    if (s.sctFunction.trim()) payload['sct_function'] = s.sctFunction.trim();
    if (s.urgencyReason.trim()) payload['sct_urgency_reason'] = s.urgencyReason.trim();
    if (s.emergencyDeclaredBy.trim()) payload['sct_emergency_declared_by_email'] = s.emergencyDeclaredBy.trim();
    if (s.deliverTo.trim()) payload['sct_delivery_address'] = s.deliverTo.trim();
    if (s.equipmentAssetTag.trim()) payload['sct_equipment_asset_tag'] = s.equipmentAssetTag.trim();
    if (s.customerName.trim()) payload['sct_customer_name'] = s.customerName.trim();
    if (s.autoCostCentre) payload['sct_cost_center'] = s.autoCostCentre;

    // ── Choice fields → Edm.Int32, base 854000000 (confirmed from Dataverse record) ──
    const buInt = choiceInt(BUSINESS_UNIT_INT, s.businessUnit);
    if (buInt !== undefined) payload['sct_business_unit'] = buInt;

    const ciInt = choiceInt(COST_IS_FOR_INT, s.costIsFor);
    if (ciInt !== undefined) payload['sct_cost_is_for'] = ciInt;

    const ptInt = choiceInt(PURCHASE_TYPE_INT, s.purchaseType);
    if (ptInt !== undefined) payload['sct_purchase_type'] = ptInt;

    const pbInt = choiceInt(PRICING_BASIS_INT, s.pricingBasis);
    if (pbInt !== undefined) payload['sct_pricing_basis'] = pbInt;

    const ugInt = choiceInt(URGENCY_INT, s.urgency);
    if (ugInt !== undefined) payload['sct_urgency'] = ugInt;

    // ── Currency / numeric ────────────────────────────────────────────────────
    const amt = parseFloat(s.estimatedAmount);
    if (!isNaN(amt)) payload['sct_estimated_total_value'] = amt;

    // ── Dates — Edm.Date fields require YYYY-MM-DD only, no time/timezone component ──
    const termStart = dateOnly(s.termStartDate);
    const termEnd = dateOnly(s.termEndDate);
    const needed = dateOnly(s.dateNeededBy);
    if (termStart) payload['sct_term_start_date'] = termStart;
    if (termEnd) payload['sct_term_end_date'] = termEnd;
    if (needed) payload['sct_date_needed_by'] = needed;

    // ── Booleans — names confirmed from actual Dataverse record ───────────────
    payload['sct_is_new_supplier'] = s.isNewSupplier;
    payload['sct_cost_recharged_to_customer'] = s.costRechargedToCustomer;
    payload['sct_customer_agreement_attached'] = s.customerAgreementAttached;
    payload['sct_tech_or_data_involved'] = s.techOrDataInvolved;
    payload['sct_supplier_on_premises'] = s.supplierOnPremises;
    payload['sct_possible_capital_purchase'] = s.possibleCapitalPurchase;
    payload['sct_is_confidential'] = s.isConfidential;
    payload['sct_is_retrospective'] = s.isRetrospective;
    payload['sct_declaration_no_prior_commitment'] = s.declarationNoPriorCommitment;
    payload['sct_declaration_no_conflict_of_interest'] = s.declarationNoConflictOfInterest;

    // ── Lookup nav-prop bindings ──────────────────────────────────────────────────
    //
    // Rules for every @odata.bind property:
    //   1. Only include if the GUID is a valid, non-empty GUID string.
    //   2. Never send a null/undefined value — omit the key entirely.
    //   3. Never send the bare logical name (lowercase) as the property key —
    //      always use the schema/nav-prop name so Dataverse doesn't 400.
    //
    const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

    // Requester — sct_Requester nav prop (schema name 'sct_Requester')
    if (s.requesterGuid && GUID_RE.test(s.requesterGuid))
      payload['sct_Requester@odata.bind'] = `/systemusers(${s.requesterGuid})`;

    // Site — sct_Site nav prop
    if (s.selectedSiteId && GUID_RE.test(s.selectedSiteId))
      payload['sct_Site@odata.bind'] = `/sct_sites_masters(${s.selectedSiteId})`;

    // Suggested Supplier — sct_Suggested_Supplier nav prop
    if (!s.isNewSupplier && s.selectedSupplierId && GUID_RE.test(s.selectedSupplierId))
      payload['sct_Suggested_Supplier@odata.bind'] = `/sct_suppliers(${s.selectedSupplierId})`;

    // Manager — automatically set manager lookup field using cr2a4_manager@odata.bind
    if (s.managerGuid && GUID_RE.test(s.managerGuid)) {
      payload['cr2a4_manager@odata.bind'] = `/systemusers(${s.managerGuid})`;
      const navProp = (s.managerNavProp && s.managerNavProp.toLowerCase() !== 'sct_manager')
        ? s.managerNavProp
        : 'cr2a4_Manager';
      if (navProp !== 'cr2a4_manager') {
        payload[`${navProp}@odata.bind`] = `/systemusers(${s.managerGuid})`;
      }
    }

    // Canonical status & pending tier:
    // Draft = 854000001 ('Draft'); Submitted = 854000002 ('Submitted')
    payload['sct_request_status'] = isDraft ? 854000001 : 854000002;
    payload['cr2a4_pendingtier'] = isDraft ? null : 'Manager';

    return payload;
  }

  // ── Submit ────────────────────────────────────────────────────────────────────

  private async _submit(isDraft: boolean): Promise<void> {
    if (!isDraft && !this._validate()) return;
    if (!this._service) return;

    const key = isDraft ? 'savingDraft' : 'submitting';
    this.setState({ [key]: true, isSubmitting: true, submitError: null } as unknown as INewRequestFormState);

    try {
      const payload = this._buildPayload(isDraft);
      let newId: string;
      const isExisting = Boolean(this.state.savedId && GUID_RE.test(this.state.savedId));

      if (isExisting) {
        newId = this.state.savedId!;
        await this._service.updateRequest(newId, payload);
      } else {
        try {
          newId = await this._service.createRequest(payload);
        } catch (submitErr: unknown) {
          const errMsg = submitErr instanceof Error ? submitErr.message : String(submitErr);
          if (errMsg.indexOf('manager') !== -1 || errMsg.indexOf('cr2a4_Manager') !== -1 || errMsg.indexOf('sct_Manager') !== -1) {
            console.warn('[NewRequestForm] Retrying request creation without manager lookup bind:', errMsg);
            const cleanPayload = { ...payload };
            Object.keys(cleanPayload).forEach(k => {
              if (k.toLowerCase().indexOf('manager@odata.bind') !== -1) {
                delete cleanPayload[k];
              }
            });
            newId = await this._service.createRequest(cleanPayload);
          } else {
            throw submitErr;
          }
        }
      }

      // ── Initial audit log: records that the request was submitted to Manager ──
      if (!isDraft) {
        const auditPayload: Record<string, unknown> = {
          sct_action: isExisting ? 2 : 1, // 2 = Edited/Resubmitted, 1 = Submitted
          sct_description: isExisting
            ? 'Request revised and resubmitted for Line Manager approval'
            : 'Request submitted for Line Manager approval',
          sct_oldvalue: 'Status: Draft, Pending Tier: None',
          sct_newvalue: 'Status: Submitted, Pending Tier: Manager',
          'sct_requestid@odata.bind': `/sct_requests(${newId})`,
        };
        if (this.state.requesterGuid && GUID_RE.test(this.state.requesterGuid)) {
          auditPayload['sct_actionby@odata.bind'] = `/systemusers(${this.state.requesterGuid})`;
        }
        try {
          await this._service.createAuditLog(auditPayload);
        } catch (auditErr) {
          console.warn('[NewRequestForm] Primary audit log failed, retrying without actionby:', auditErr);
          delete auditPayload['sct_actionby@odata.bind'];
          await this._service.createAuditLog(auditPayload).catch(() => { /* non-fatal */ });
        }
      }

      if (this.state.attachments.length > 0) {
        this.setState({ uploadingFiles: true });
        await Promise.all(
          this.state.attachments.map(a =>
            this._service!.uploadAnnotation(newId, a.name, a.mimeType, a.base64)
          )
        );
        this.setState({ uploadingFiles: false });
      }

      const submitMessage = isDraft
        ? 'Draft saved. You can retrieve it from your Requests.'
        : 'Request submitted successfully! It is now pending Manager approval.';

      this.setState({
        [key]: false,
        isSubmitting: false,
        submitSuccess: submitMessage,
        savedId: newId,
      } as unknown as INewRequestFormState);

      // ── Send approval email to Line Manager (resolved from Dataverse systemusers) ──
      if (!isDraft) {
        void (async () => {
          try {
            const graphClient: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient('3');
            let managerEmail = '';

            // 1. Primary: Resolve manager from Dataverse systemusers using managerGuid
            if (this.state.managerGuid && this._service) {
              const mgr = await this._service.getSystemUserById(this.state.managerGuid).catch(() => null)
                ?? await this._service.getCurrentSystemUser(this.state.managerGuid).catch(() => null)
                ?? await this._service.getSystemUserByEmail(this.state.managerGuid).catch(() => null);
              if (mgr && mgr.internalemailaddress) {
                managerEmail = mgr.internalemailaddress;
              }
            }

            // 2. Fallback: Graph /me/manager if Dataverse record did not carry email
            if (!managerEmail) {
              const mgrGraph = await getManager(graphClient).catch(() => null);
              if (mgrGraph && mgrGraph.mail) {
                managerEmail = mgrGraph.mail;
              }
            }

            // Alert Manager upon successful submission: to: [ManagerEmail], cc: [Current_User_Email]
            const currentUserEmail = this.props.context.pageContext.user.email || this.state.requesterEmail || '';
            const toEmails = managerEmail ? [managerEmail] : [];
            const ccEmails = currentUserEmail ? [currentUserEmail] : [];

            // Attempt to fetch newly created record to read canonical autonumber reference (e.g. SCT-2026-1051)
            let createdRec: ODataRecord | null = null;
            if (this._service) {
              try {
                createdRec = await this._service.getRawRequestById(newId);
              } catch { /* non-fatal */ }
            }

            const refNumber = createdRec
              ? (strVal(createdRec, 'sct_sct_reference') || strVal(createdRec, 'sct_referencenumber') || newId)
              : '(New Submission)';

            const subject = `[SCT Approval Required] ${refNumber} — Line Manager Approval`;
            const siteUrl = this.props.context.pageContext.web.absoluteUrl;

            const requestSummary: Record<string, unknown> = createdRec || {
              sct_sct_reference: refNumber,
              sct_referencenumber: newId,
              sct_request_title: this.state.requestTitle,
              sct_estimated_total_value: this.state.estimatedAmount,
              _sct_requester_value: this.state.requesterName || this.props.context.pageContext.user.displayName,
              sct_business_unit: this.state.businessUnit,
              _sct_site_value: this.state.selectedSiteName || this.state.selectedSiteId,
              sct_description: this.state.description,
            };

            const htmlBody = generateApprovalEmailHTML(
              requestSummary,
              [],
              siteUrl,
              'Line Manager Review & Approval Required'
            );

            if (toEmails.length > 0) {
              await sendGraphEmail(graphClient, toEmails, ccEmails, subject, htmlBody);
            }
          } catch (mailErr) {
            console.warn('[NewRequestForm] Manager notification email dispatch warning:', mailErr);
          }
        })();
      }
    } catch (err: unknown) {
      this.setState({
        [key]: false,
        isSubmitting: false,
        submitError: err instanceof Error ? `Network error: ${err.message}` : 'Network error: Please try again.',
      } as unknown as INewRequestFormState);
    }
  }

  // ── Attachment handling ───────────────────────────────────────────────────────

  private async _onFilesSelected(files: FileList | null): Promise<void> {
    if (!files || files.length === 0) return;
    const added: IAttachment[] = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      try {
        const base64 = await readFileAsBase64(file);
        added.push({ name: file.name, mimeType: file.type || 'application/octet-stream', base64, sizeKb: Math.round(file.size / 1024) });
      } catch { /* skip unreadable */ }
    }
    this.setState(prev => ({ attachments: [...prev.attachments, ...added] }));
    if (this._fileInputRef.current) this._fileInputRef.current.value = '';
  }

  private _removeAttachment(index: number): void {
    this.setState(prev => ({ attachments: prev.attachments.filter((_, i) => i !== index) }));
  }

  // ── Render helpers ────────────────────────────────────────────────────────────

  private _err(field: string): string | undefined {
    return this.state.errors[field] || undefined;
  }

  private _clrErr(field: string): void {
    if (this.state.errors[field]) {
      const errs = { ...this.state.errors };
      delete errs[field];
      this.setState({ errors: errs });
    }
  }

  // ── Section 1: Header & Requester ─────────────────────────────────────────────

  private _renderHeaderSection(): React.ReactElement {
    const { sites, sitesError, selectedSiteId, businessUnit, sctFunction, costIsFor, autoCostCentre, autoDivision, autoBudgetOwner } = this.state;
    // Build dropdown options from raw ODataRecord[] — field names resolved at runtime
    const siteOptions: IDropdownOption[] = sites
      .map(s => ({ key: this._siteId(s), text: this._siteLabel(s) }))
      .filter(o => o.key);
    const showMultiSiteNote = costIsFor === 'Several Sites' || costIsFor === 'Company-Wide';
    const { requesterGuid, requesterName, requesterEmail, userLookupError, loadingMeta } = this.state;

    return (
      <Section iconName="Contact" iconBg="#EBF3FB" iconColor="#0078D4" title="Requester & Site Information">
        {sitesError && (
          <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { borderRadius: 6, marginBottom: 12 } }}>
            Could not load sites: {sitesError}. Check Dataverse permissions and entity set name.
          </MessageBar>
        )}
        {userLookupError && (
          <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { borderRadius: 6, marginBottom: 12 } }}>
            {userLookupError}
          </MessageBar>
        )}

        {/* Read-only requester strip — resolved from Dataverse systemuser */}
        <div style={{ display: 'flex', gap: 16, marginBottom: 16, background: '#F8F8F8', borderRadius: 8, padding: '12px 16px', border: '1px solid #EDEBE9', alignItems: 'center' }}>
          <div style={{ flex: 1 }}>
            <Text variant="small" style={{ color: '#A19F9D', display: 'block', marginBottom: 2 }}>REQUESTER</Text>
            {loadingMeta
              ? <Text variant="small" style={{ color: '#605E5C' }}>Resolving user…</Text>
              : <>
                <Text variant="medium" style={{ color: '#201F1E', fontWeight: 600, display: 'block' }}>
                  {requesterName || this.props.context.pageContext.user.displayName}
                </Text>
                {requesterEmail && (
                  <Text variant="small" style={{ color: '#605E5C' }}>{requesterEmail}</Text>
                )}
              </>
            }
          </div>
          <div style={{ flex: 1 }}>
            <Text variant="small" style={{ color: '#A19F9D', display: 'block', marginBottom: 2 }}>DATAVERSE BINDING</Text>
            {requesterGuid
              ? <span style={{ background: '#DFF6DD', color: '#107C10', fontSize: 11, fontWeight: 700, padding: '3px 10px', borderRadius: 12 }}>
                Resolved ✓
              </span>
              : <span style={{ background: '#FFF4CE', color: '#905400', fontSize: 11, fontWeight: 700, padding: '3px 10px', borderRadius: 12 }}>
                {loadingMeta ? 'Resolving…' : 'Not bound'}
              </span>
            }
          </div>
          <div style={{ flex: 1 }}>
            <Text variant="small" style={{ color: '#A19F9D', display: 'block', marginBottom: 2 }}>SCT REFERENCE</Text>
            <span style={{ background: '#EDEBE9', color: '#605E5C', fontSize: 12, fontWeight: 600, padding: '3px 12px', borderRadius: 12 }}>
              Draft — assigned on submission
            </span>
          </div>
        </div>

        {/* Read-only Line Manager strip / field — fetched from Azure Active Directory */}
        <Row>
          <TextField
            label="Line Manager"
            value={this.state.managerName ? `${this.state.managerName}${this.state.managerEmail ? ` (${this.state.managerEmail})` : ''}` : (this.state.loadingMeta ? 'Fetching manager from Azure AD…' : 'No manager assigned')}
            readOnly
            disabled
            description="Automatically fetched from Azure Active Directory (read-only)"
            styles={fieldStyles}
          />
        </Row>

        <Row>
          <Dropdown
            label="Business Unit"
            required
            placeholder="Select business unit"
            options={BUSINESS_UNIT_OPTIONS}
            selectedKey={businessUnit || undefined}
            onChange={(_, opt) => { this.setState({ businessUnit: String(opt?.key ?? '') }); this._clrErr('businessUnit'); }}
            errorMessage={this._err('businessUnit')}
            styles={dropdownStyles}
          />
          <Dropdown
            label="Site / Location"
            required
            placeholder="Select site"
            options={siteOptions}
            selectedKey={selectedSiteId || undefined}
            onChange={(_, opt) => { this._onSiteSelected(String(opt?.key ?? '')); this._clrErr('selectedSiteId'); }}
            errorMessage={this._err('selectedSiteId')}
            styles={dropdownStyles}
          />
        </Row>

        {/* Auto-filled site-derived fields — read only */}
        {(autoCostCentre || autoDivision || autoBudgetOwner) && (
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 24, background: '#F0F6FF', borderRadius: 8, padding: '12px 16px', marginBottom: 16, border: '1px solid #C7E0F4' }}>
            {autoCostCentre && (
              <div>
                <Text variant="small" style={{ color: '#0078D4', display: 'block', fontWeight: 600 }}>Cost Centre</Text>
                <Text variant="medium" style={{ color: '#201F1E' }}>{autoCostCentre}</Text>
              </div>
            )}
            {autoDivision && (
              <div>
                <Text variant="small" style={{ color: '#0078D4', display: 'block', fontWeight: 600 }}>Division / Region</Text>
                <Text variant="medium" style={{ color: '#201F1E' }}>{autoDivision}</Text>
              </div>
            )}
            {autoBudgetOwner && (
              <div>
                <Text variant="small" style={{ color: '#0078D4', display: 'block', fontWeight: 600 }}>L1 Budget Owner</Text>
                <Text variant="medium" style={{ color: '#201F1E' }}>{autoBudgetOwner}</Text>
              </div>
            )}
            <div style={{ alignSelf: 'center' }}>
              <Text variant="small" style={{ color: '#605E5C' }}>
                <Icon iconName="Lock" style={{ fontSize: 11, marginRight: 4 }} />
                Auto-filled from Site Master — not editable by requester
              </Text>
            </div>
          </div>
        )}

        <Row>
          <TextField
            label="Function"
            placeholder="e.g. Operations, IT, Finance"
            value={sctFunction}
            onChange={(_, v) => this.setState({ sctFunction: v ?? '' })}
            styles={fieldStyles}
          />
          <Dropdown
            label="Cost Is For"
            placeholder="Select scope"
            options={COST_IS_FOR_OPTIONS}
            selectedKey={costIsFor || undefined}
            onChange={(_, opt) => this.setState({ costIsFor: String(opt?.key ?? '') })}
            styles={dropdownStyles}
          />
        </Row>

        {showMultiSiteNote && (
          <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: 6, marginBottom: 16 } }}>
            This request spans multiple sites or business units. Attach a cost allocation breakdown with your supporting documents.
          </MessageBar>
        )}
      </Section>
    );
  }

  // ── Section 2: Requirement & Supplier ─────────────────────────────────────────

  private _renderRequirementSection(): React.ReactElement {
    const { requestTitle, description, purchaseType, isNewSupplier, selectedSupplierId,
      newSupplierName, newSupplierLegalName, newSupplierContactName,
      newSupplierContactEmail, newSupplierContactPhone, newSupplierW9Notes, suppliers } = this.state;

    const supplierOptions: IDropdownOption[] = suppliers.map(s => ({
      key: s.sct_supplierid,
      text: s.sct_supplier_name || s.sct_suppliername || s.sct_supplierid,
    }));

    const isRepairOrRental = purchaseType === 'Repair' || purchaseType === 'Rental';
    const isTempLabor = purchaseType === 'Temporary Labor';

    return (
      <Section iconName="BuildDefinition" iconBg="#F4F0FF" iconColor="#8764B8" title="Requirement & Supplier">
        <TextField
          label="Request Title"
          required
          placeholder="Brief, specific title for this spend request"
          value={requestTitle}
          onChange={(_, v) => { this.setState({ requestTitle: v ?? '' }); this._clrErr('requestTitle'); }}
          errorMessage={this._err('requestTitle')}
          styles={fieldStyles}
        />
        <Row>
          <Dropdown
            label="Purchase Type"
            required
            placeholder="Select type"
            options={PURCHASE_TYPE_OPTIONS}
            selectedKey={purchaseType || undefined}
            onChange={(_, opt) => { this.setState({ purchaseType: String(opt?.key ?? '') }); this._clrErr('purchaseType'); }}
            errorMessage={this._err('purchaseType')}
            styles={dropdownStyles}
          />
          <div />
        </Row>

        {isTempLabor && (
          <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { borderRadius: 6, marginBottom: 16 } }}>
            Temporary Labour requests at a unionized site require prior union compliance review. Confirm with HR / Site Manager before submission.
          </MessageBar>
        )}
        {isRepairOrRental && (
          <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: 6, marginBottom: 16 } }}>
            Repair and Rental requests require an equipment asset tag or serial number in Section 4.
          </MessageBar>
        )}

        <TextField
          label="Description / Justification"
          required
          placeholder="Describe what is being purchased, why it is needed, and expected business outcome"
          multiline
          rows={4}
          value={description}
          onChange={(_, v) => { this.setState({ description: v ?? '' }); this._clrErr('description'); }}
          errorMessage={this._err('description')}
          styles={fieldStyles}
        />

        {/* ── Supplier toggle ── */}
        <Toggle
          label="Is this a new / proposed supplier?"
          checked={isNewSupplier}
          onChange={(_, checked) => this.setState({ isNewSupplier: !!checked, selectedSupplierId: '', newSupplierName: '' })}
          onText="Yes — proposed / new supplier"
          offText="No — select from approved supplier list"
          styles={{ root: { marginBottom: 16 } }}
        />

        {!isNewSupplier ? (
          <Dropdown
            label="Suggested Supplier"
            placeholder="Select an approved supplier"
            options={supplierOptions}
            selectedKey={selectedSupplierId || undefined}
            onChange={(_, opt) => this.setState({ selectedSupplierId: String(opt?.key ?? '') })}
            styles={dropdownStyles}
          />
        ) : (
          <div style={{ background: '#FFF4CE', borderRadius: 8, padding: '12px 16px', marginBottom: 16, border: '1px solid #FFB900' }}>
            <Text variant="small" style={{ color: '#7A4800', display: 'block', marginBottom: 12, fontWeight: 600 }}>
              <Icon iconName="Warning" style={{ marginRight: 6 }} />
              New supplier details below do not automatically make this supplier Active or Preferred. Procurement review and approval is required before any commitment can be made.
            </Text>
            <TextField
              label="Proposed Supplier Name" required
              placeholder="Trading / display name"
              value={newSupplierName}
              onChange={(_, v) => { this.setState({ newSupplierName: v ?? '' }); this._clrErr('newSupplierName'); }}
              errorMessage={this._err('newSupplierName')}
              styles={fieldStyles}
            />
            {(() => {
              const trimmed = (newSupplierName || '').trim().toLowerCase();
              if (trimmed.length >= 3) {
                const dup = this.state.suppliers.find(sup => {
                  const sName = (sup.sct_supplier_name || sup.sct_suppliername || '').toLowerCase();
                  return Boolean(sName && (sName.indexOf(trimmed) !== -1 || trimmed.indexOf(sName) !== -1));
                });
                if (dup) {
                  const matchName = dup.sct_supplier_name || dup.sct_suppliername || 'Existing Vendor';
                  return (
                    <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { borderRadius: 6, margin: '8px 0' } }}>
                      Approved vendor <strong>{matchName}</strong> already exists in the system. Verify if this request can use the approved vendor.
                    </MessageBar>
                  );
                }
              }
              return null;
            })()}
            <TextField
              label="Legal / Registered Name"
              placeholder="Full registered legal entity name"
              value={newSupplierLegalName}
              onChange={(_, v) => this.setState({ newSupplierLegalName: v ?? '' })}
              styles={fieldStyles}
            />
            <Row>
              <TextField
                label="Contact Name"
                placeholder="Primary contact person"
                value={newSupplierContactName}
                onChange={(_, v) => this.setState({ newSupplierContactName: v ?? '' })}
                styles={fieldStyles}
              />
              <TextField
                label="Contact Email"
                placeholder="contact@supplier.com"
                type="email"
                value={newSupplierContactEmail}
                onChange={(_, v) => this.setState({ newSupplierContactEmail: v ?? '' })}
                styles={fieldStyles}
              />
            </Row>
            <Row>
              <TextField
                label="Contact Phone"
                placeholder="+61 ..."
                value={newSupplierContactPhone}
                onChange={(_, v) => this.setState({ newSupplierContactPhone: v ?? '' })}
                styles={fieldStyles}
              />
              <TextField
                label="W-9 / Bank Details Notes"
                placeholder="Note status of tax form and bank details verification"
                value={newSupplierW9Notes}
                onChange={(_, v) => this.setState({ newSupplierW9Notes: v ?? '' })}
                styles={fieldStyles}
              />
            </Row>
          </div>
        )}
      </Section>
    );
  }

  // ── Section 3: Financials, Term, Urgency ──────────────────────────────────────

  private _renderFinancialSection(): React.ReactElement {
    const { estimatedAmount, termStartDate, termEndDate, pricingBasis,
      dateNeededBy, urgency, urgencyReason, emergencyDeclaredBy, isNewSupplier } = this.state;

    const amt = parseFloat(estimatedAmount) || 0;
    const guidance = calcApprovalGuidance(isNewSupplier, amt);
    const showUrgencyReason = urgency === 'Urgent' || urgency === 'Emergency';
    const isEmergency = urgency === 'Emergency';

    return (
      <Section iconName="Money" iconBg="#DFF6DD" iconColor="#107C10" title="Financials, Term & Urgency">

        {/* Pre-approval guidance banner */}
        {(amt > 0 || isNewSupplier) && (
          <div style={{ background: guidance.bg, borderRadius: 8, padding: '12px 16px', marginBottom: 16, border: `1px solid ${guidance.color}40` }}>
            <Text variant="small" style={{ color: guidance.color, fontWeight: 700, display: 'block', marginBottom: 6 }}>
              <Icon iconName={guidance.iconName} style={{ marginRight: 6, fontSize: 13 }} />
              Pre-Approval Evidence Required
            </Text>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              {guidance.levels.map(l => (
                <span key={l} style={{ background: 'rgba(255,255,255,0.7)', color: guidance.color, fontSize: 12, fontWeight: 600, padding: '3px 10px', borderRadius: 10, border: `1px solid ${guidance.color}50` }}>
                  {l}
                </span>
              ))}
            </div>
            <Text variant="small" style={{ color: guidance.color, marginTop: 6, display: 'block' }}>
              Written pre-approval from the above level{guidance.levels.length > 1 ? 's' : ''} must be attached before submission.
            </Text>
          </div>
        )}

        <Row>
          <TextField
            label="Estimated Total Value (AUD)"
            required
            placeholder="0.00"
            prefix="$"
            type="number"
            value={estimatedAmount}
            onChange={(_, v) => { this.setState({ estimatedAmount: v ?? '' }); this._clrErr('estimatedAmount'); }}
            errorMessage={this._err('estimatedAmount')}
            description="Full value over the entire intended term or contract period"
            styles={fieldStyles}
          />
          <Dropdown
            label="Pricing Basis"
            placeholder="Select pricing basis"
            options={PRICING_BASIS_OPTIONS}
            selectedKey={pricingBasis || undefined}
            onChange={(_, opt) => this.setState({ pricingBasis: String(opt?.key ?? '') })}
            styles={dropdownStyles}
          />
        </Row>

        <Row>
          <DatePicker
            label="Term Start Date"
            placeholder="Select start date"
            firstDayOfWeek={DayOfWeek.Monday}
            value={termStartDate ?? undefined}
            onSelectDate={d => this.setState({ termStartDate: d ?? null })}
            styles={{ root: { marginBottom: 16 } }}
          />
          <DatePicker
            label="Term End Date"
            placeholder="Select end date"
            firstDayOfWeek={DayOfWeek.Monday}
            value={termEndDate ?? undefined}
            onSelectDate={d => this.setState({ termEndDate: d ?? null })}
            styles={{ root: { marginBottom: 16 } }}
          />
        </Row>

        <Row>
          <DatePicker
            label="Date Needed By"
            isRequired
            placeholder="Select required date"
            firstDayOfWeek={DayOfWeek.Monday}
            value={dateNeededBy ?? undefined}
            onSelectDate={d => { this.setState({ dateNeededBy: d ?? null }); this._clrErr('dateNeededBy'); }}
            styles={{ root: { marginBottom: 16 } }}
          />
          <Dropdown
            label="Urgency"
            required
            placeholder="Select urgency"
            options={URGENCY_OPTIONS}
            selectedKey={urgency || undefined}
            onChange={(_, opt) => { this.setState({ urgency: String(opt?.key ?? '') }); this._clrErr('urgency'); }}
            errorMessage={this._err('urgency')}
            styles={dropdownStyles}
          />
        </Row>

        {showUrgencyReason && (
          <TextField
            label="Urgency Reason"
            required
            placeholder="Explain why this request is urgent or an emergency"
            multiline rows={3}
            value={urgencyReason}
            onChange={(_, v) => { this.setState({ urgencyReason: v ?? '' }); this._clrErr('urgencyReason'); }}
            errorMessage={this._err('urgencyReason')}
            styles={fieldStyles}
          />
        )}

        {isEmergency && (
          <MessageBar messageBarType={MessageBarType.severeWarning} styles={{ root: { borderRadius: 6, marginBottom: 16 } }}>
            <strong>Emergency Purchase Cap:</strong> Emergency spend is limited to <strong>$25,000 AUD and 7 days duration</strong> unless an approved written override from CFO is attached. The person declaring the emergency must be recorded below.
          </MessageBar>
        )}

        {isEmergency && (
          <TextField
            label="Emergency Declared By"
            required
            placeholder="Full name and title of person authorising this emergency"
            value={emergencyDeclaredBy}
            onChange={(_, v) => { this.setState({ emergencyDeclaredBy: v ?? '' }); this._clrErr('emergencyDeclaredBy'); }}
            errorMessage={this._err('emergencyDeclaredBy')}
            styles={fieldStyles}
          />
        )}
      </Section>
    );
  }

  // ── Section 4: Logistics & Conditional Triggers ───────────────────────────────

  private _renderLogisticsSection(): React.ReactElement {
    const { deliverTo, receivingContact, equipmentAssetTag, purchaseType,
      costRechargedToCustomer, customerName, customerAgreementAttached,
      techOrDataInvolved, supplierOnPremises, possibleCapitalPurchase,
      isConfidential, isRetrospective } = this.state;

    const isRepairOrRental = purchaseType === 'Repair' || purchaseType === 'Rental';

    return (
      <Section iconName="DeliveryTruck" iconBg="#FFF4CE" iconColor="#905400" title="Logistics & Conditions">
        <Row>
          <TextField
            label="Deliver To / Delivery Address"
            placeholder="Site address or internal delivery location"
            multiline rows={2}
            value={deliverTo}
            onChange={(_, v) => this.setState({ deliverTo: v ?? '' })}
            styles={fieldStyles}
          />
          <TextField
            label="Receiving Contact"
            placeholder="Name of person receiving goods / services"
            value={receivingContact}
            onChange={(_, v) => this.setState({ receivingContact: v ?? '' })}
            styles={fieldStyles}
          />
        </Row>

        <TextField
          label={`Equipment Asset Tag / Serial Number${isRepairOrRental ? ' *' : ''}`}
          placeholder="Asset tag or serial number for the item being repaired or rented"
          value={equipmentAssetTag}
          onChange={(_, v) => { this.setState({ equipmentAssetTag: v ?? '' }); this._clrErr('equipmentAssetTag'); }}
          errorMessage={this._err('equipmentAssetTag')}
          description={isRepairOrRental ? 'Required for Repair and Rental purchase types' : 'If applicable'}
          styles={fieldStyles}
        />

        {/* Conditional flags grid */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '0 24px', marginTop: 8 }}>
          <Toggle
            label="Cost Recharged to Customer?"
            checked={costRechargedToCustomer}
            onChange={(_, c) => this.setState({ costRechargedToCustomer: !!c })}
            onText="Yes" offText="No"
            styles={{ root: { marginBottom: 16 } }}
          />
          <Toggle
            label="Technology or Data Involved?"
            checked={techOrDataInvolved}
            onChange={(_, c) => this.setState({ techOrDataInvolved: !!c })}
            onText="Yes — CIO approval required (Rule R17)" offText="No"
            styles={{ root: { marginBottom: 16 } }}
          />
          <Toggle
            label="Supplier Working on Premises?"
            checked={supplierOnPremises}
            onChange={(_, c) => this.setState({ supplierOnPremises: !!c })}
            onText="Yes — Risk COI check required" offText="No"
            styles={{ root: { marginBottom: 16 } }}
          />
          <Toggle
            label="Possible Capital Purchase?"
            checked={possibleCapitalPurchase}
            onChange={(_, c) => this.setState({ possibleCapitalPurchase: !!c })}
            onText="Yes — Finance capital review triggered" offText="No"
            styles={{ root: { marginBottom: 16 } }}
          />
          <Toggle
            label="Confidential / Legal / HR Matter?"
            checked={isConfidential}
            onChange={(_, c) => this.setState({ isConfidential: !!c })}
            onText="Yes — item-level access restriction applied" offText="No"
            styles={{ root: { marginBottom: 16 } }}
          />
          <Toggle
            label="Retrospective Request?"
            checked={isRetrospective}
            onChange={(_, c) => this.setState({ isRetrospective: !!c })}
            onText="Yes — commitment may precede approval" offText="No"
            styles={{ root: { marginBottom: 16 } }}
          />
        </div>

        {/* Cost recharge conditional fields */}
        {costRechargedToCustomer && (
          <div style={{ background: '#EBF3FB', borderRadius: 8, padding: '12px 16px', marginTop: 8, border: '1px solid #C7E0F4' }}>
            <Row>
              <TextField
                label="Customer Name" required
                placeholder="Customer or project being recharged"
                value={customerName}
                onChange={(_, v) => { this.setState({ customerName: v ?? '' }); this._clrErr('customerName'); }}
                errorMessage={this._err('customerName')}
                styles={fieldStyles}
              />
              <div style={{ paddingTop: 36 }}>
                <Checkbox
                  label="Customer agreement / quote attached"
                  checked={customerAgreementAttached}
                  onChange={(_, c) => { this.setState({ customerAgreementAttached: !!c }); this._clrErr('customerAgreementAttached'); }}
                  styles={checkStyles}
                />
                {this._err('customerAgreementAttached') && (
                  <Text variant="small" style={{ color: '#A4262C', display: 'block', marginTop: 4 }}>
                    {this._err('customerAgreementAttached')}
                  </Text>
                )}
              </div>
            </Row>
          </div>
        )}

        {/* Info banners for triggered reviews */}
        {techOrDataInvolved && (
          <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: 6, marginTop: 12 } }}>
            CIO approval (Rule R17) has been triggered. The SCT desk will route this request accordingly after submission.
          </MessageBar>
        )}
        {supplierOnPremises && (
          <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { borderRadius: 6, marginTop: 8 } }}>
            Risk Management COI (Certificate of Insurance) check has been triggered. Ensure the supplier's insurance certificate is on file before work begins.
          </MessageBar>
        )}
        {possibleCapitalPurchase && (
          <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: 6, marginTop: 8 } }}>
            Finance capital review has been triggered. FP&A will assess whether this spend should be capitalised.
          </MessageBar>
        )}
        {isRetrospective && (
          <MessageBar messageBarType={MessageBarType.severeWarning} styles={{ root: { borderRadius: 6, marginTop: 8 } }}>
            Retrospective requests indicate a commitment may have been made before SCT approval. This requires escalation and ratification review.
          </MessageBar>
        )}
      </Section>
    );
  }

  // ── Section 5: Declarations ────────────────────────────────────────────────────

  private _renderDeclarationsSection(): React.ReactElement {
    const { declarationNoPriorCommitment, declarationNoConflictOfInterest } = this.state;
    return (
      <Section iconName="Shield" iconBg="#FDE7E9" iconColor="#A4262C" title="Mandatory Declarations">
        <Text variant="small" style={{ color: '#605E5C', display: 'block', marginBottom: 16, lineHeight: '1.6' }}>
          Both declarations must be confirmed before this request can be submitted. False declarations may result in disciplinary action.
        </Text>
        <div style={{ background: '#FDE7E9', borderRadius: 8, padding: '16px 20px', border: '1px solid #F4B8BC' }}>
          <Checkbox
            label="Rule R1 — No Prior Commitment: I confirm that no commitment, agreement, or purchase order has been made or placed with any supplier prior to obtaining the necessary SCT approvals."
            checked={declarationNoPriorCommitment}
            onChange={(_, c) => { this.setState({ declarationNoPriorCommitment: !!c }); this._clrErr('declarationNoPriorCommitment'); }}
            styles={{ root: { marginBottom: 14 }, label: { color: '#201F1E', lineHeight: '1.5' } }}
          />
          {this._err('declarationNoPriorCommitment') && (
            <Text variant="small" style={{ color: '#A4262C', marginBottom: 10, display: 'block' }}>
              {this._err('declarationNoPriorCommitment')}
            </Text>
          )}
          <Checkbox
            label="Rule R10 — No Conflict of Interest: I declare that I have no personal, financial, or professional conflict of interest with any supplier relevant to this request."
            checked={declarationNoConflictOfInterest}
            onChange={(_, c) => { this.setState({ declarationNoConflictOfInterest: !!c }); this._clrErr('declarationNoConflictOfInterest'); }}
            styles={{ root: { marginBottom: 0 }, label: { color: '#201F1E', lineHeight: '1.5' } }}
          />
          {this._err('declarationNoConflictOfInterest') && (
            <Text variant="small" style={{ color: '#A4262C', marginTop: 4, display: 'block' }}>
              {this._err('declarationNoConflictOfInterest')}
            </Text>
          )}
        </div>
      </Section>
    );
  }

  // ── Section 6: Attachments ────────────────────────────────────────────────────

  private _renderAttachmentsSection(): React.ReactElement {
    const { attachments, isNewSupplier, estimatedAmount } = this.state;
    const amt = parseFloat(estimatedAmount) || 0;
    const guidance = calcApprovalGuidance(isNewSupplier, amt);

    return (
      <Section iconName="Attach" iconBg="#E6F7F4" iconColor="#00B294" title="Supporting Documents & Evidence">
        {(amt > 0 || isNewSupplier) && (
          <div style={{ background: guidance.bg, borderRadius: 6, padding: '10px 14px', marginBottom: 14, border: `1px solid ${guidance.color}40` }}>
            <Text variant="small" style={{ color: guidance.color, fontWeight: 600 }}>
              Required attachments for this request: {guidance.levels.join(', ')} pre-approval evidence
              {isNewSupplier ? ', plus supplier contact details and any quotes / proposals' : ''}.
            </Text>
          </div>
        )}
        <Text variant="small" style={{ color: '#605E5C', display: 'block', marginBottom: 12 }}>
          Attach quotes, supplier proposals, written approvals, customer agreements, and any other supporting evidence.
        </Text>
        <input
          type="file" multiple ref={this._fileInputRef}
          style={{ display: 'none' }}
          onChange={e => void this._onFilesSelected(e.target.files)}
          aria-hidden="true"
        />
        <DefaultButton
          text="Attach Files"
          iconProps={{ iconName: 'Attach' }}
          onClick={() => this._fileInputRef.current?.click()}
          styles={{ root: { borderRadius: 6, marginBottom: 16 } }}
        />
        {this._err('attachments') && (
          <MessageBar messageBarType={MessageBarType.error} styles={{ root: { borderRadius: 6, marginBottom: 12 } }}>
            {this._err('attachments')}
          </MessageBar>
        )}
        {attachments.length > 0 && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {attachments.map((a, i) => (
              <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 10, background: '#F8F8F8', borderRadius: 6, padding: '8px 12px', border: '1px solid #EDEBE9' }}>
                <Icon iconName="Page" style={{ color: '#0078D4', fontSize: 16, flexShrink: 0 }} />
                <div style={{ flex: 1, minWidth: 0 }}>
                  <Text variant="small" style={{ color: '#201F1E', fontWeight: 600, display: 'block', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{a.name}</Text>
                  <Text variant="small" style={{ color: '#A19F9D' }}>{a.sizeKb} KB</Text>
                </div>
                <IconButton
                  iconProps={{ iconName: 'Cancel' }} title="Remove" ariaLabel="Remove attachment"
                  onClick={() => this._removeAttachment(i)}
                  styles={{ root: { color: '#A4262C', width: 28, height: 28 } }}
                />
              </div>
            ))}
          </div>
        )}
      </Section>
    );
  }

  // ── Wizard helpers ────────────────────────────────────────────────────────────

  private _canGoNext(step: number): boolean {
    const s = this.state;
    switch (step) {
      case 0: return !!s.businessUnit && !!s.selectedSiteId;
      case 1: {
        const rawAmt = s.estimatedAmount ? String(s.estimatedAmount).trim() : '';
        const amt = parseFloat(rawAmt.replace(/[$,]/g, ''));
        const validAmt = !isNaN(amt) && amt > 0;
        const validSupplier = !s.isNewSupplier ? true : !!s.newSupplierName.trim();
        return validAmt && !!s.description.trim() && !!s.purchaseType && validSupplier && !!s.dateNeededBy && !!s.urgency;
      }
      case 2: {
        const validRecharge = !s.costRechargedToCustomer || (!!s.customerName.trim() && s.customerAgreementAttached);
        const validAsset = (s.purchaseType !== 'Repair' && s.purchaseType !== 'Rental') || !!s.equipmentAssetTag.trim();
        return validRecharge && validAsset && s.declarationNoPriorCommitment && s.declarationNoConflictOfInterest;
      }
      case 3: return true;
      default: return true;
    }
  }

  private _renderReviewStep(): React.ReactElement {
    const s = this.state;
    const field = (label: string, value: string): React.ReactElement => (
      <div style={{ display: 'flex', gap: 8, marginBottom: 10 }}>
        <Text variant="small" style={{ color: '#A19F9D', minWidth: 180, flexShrink: 0 }}>{label}</Text>
        <Text variant="small" style={{ color: '#201F1E', fontWeight: 600, wordBreak: 'break-word' }}>{value || '—'}</Text>
      </div>
    );
    const bool = (v: boolean): string => v ? 'Yes' : 'No';

    return (
      <div className={styles.card} style={{ marginBottom: 16 }}>
        <div className={styles.cardBody}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20 }}>
            <Icon iconName="DocumentSearch" style={{ fontSize: 20, color: '#0078D4' }} />
            <Text variant="large" style={{ fontWeight: 700, color: '#201F1E' }}>Review Your Request</Text>
          </div>

          {/* General Info */}
          <div style={{ marginBottom: 18 }}>
            <Text variant="mediumPlus" style={{ fontWeight: 600, color: '#0078D4', display: 'block', marginBottom: 10, borderBottom: '2px solid #EBF3FB', paddingBottom: 6 }}>
              General Info
            </Text>
            {field('Requester', s.requesterName || '—')}
            {field('Line Manager', s.managerName ? `${s.managerName}${s.managerEmail ? ` (${s.managerEmail})` : ''}` : '—')}
            {field('Business Unit', s.businessUnit)}
            {field('Site', s.selectedSiteName)}
            {field('SCT Function', s.sctFunction)}
            {field('Cost Is For', s.costIsFor)}
            {field('Purchase Type', s.purchaseType)}
          </div>

          {/* Financials */}
          <div style={{ marginBottom: 18 }}>
            <Text variant="mediumPlus" style={{ fontWeight: 600, color: '#0078D4', display: 'block', marginBottom: 10, borderBottom: '2px solid #EBF3FB', paddingBottom: 6 }}>
              Financials
            </Text>
            {field('Estimated Amount', s.estimatedAmount ? `$${parseFloat(s.estimatedAmount).toLocaleString()}` : '—')}
            {field('Cost Centre', s.autoCostCentre)}
            {field('Pricing Basis', s.pricingBasis)}
            {field('Date Needed By', s.dateNeededBy ? s.dateNeededBy.toLocaleDateString() : '—')}
            {field('Urgency', s.urgency)}
          </div>

          {/* Supplier & Compliance */}
          <div style={{ marginBottom: 18 }}>
            <Text variant="mediumPlus" style={{ fontWeight: 600, color: '#0078D4', display: 'block', marginBottom: 10, borderBottom: '2px solid #EBF3FB', paddingBottom: 6 }}>
              Supplier &amp; Compliance
            </Text>
            {field('Request Title', s.requestTitle)}
            {field('New Supplier?', bool(s.isNewSupplier))}
            {!s.isNewSupplier && field('Selected Supplier', s.selectedSupplierId ? 'Linked' : '—')}
            {s.isNewSupplier && field('New Supplier Name', s.newSupplierName)}
            {field('Tech/Data Involved?', bool(s.techOrDataInvolved))}
            {field('Supplier On Premises?', bool(s.supplierOnPremises))}
            {field('Capital Purchase?', bool(s.possibleCapitalPurchase))}
            {field('Attachments', s.attachments.length > 0 ? `${s.attachments.length} file(s)` : 'None')}
          </div>

          {/* Declarations */}
          <div>
            <Text variant="mediumPlus" style={{ fontWeight: 600, color: '#0078D4', display: 'block', marginBottom: 10, borderBottom: '2px solid #EBF3FB', paddingBottom: 6 }}>
              Declarations
            </Text>
            {field('No Prior Commitment (R1)', s.declarationNoPriorCommitment ? '✓ Confirmed' : '✗ Not confirmed')}
            {field('No Conflict of Interest (R10)', s.declarationNoConflictOfInterest ? '✓ Confirmed' : '✗ Not confirmed')}
          </div>
        </div>
      </div>
    );
  }

  // ── Success screen ─────────────────────────────────────────────────────────────

  private _renderSuccess(): React.ReactElement {
    const { savedId, submitSuccess } = this.state;
    return (
      <div className={styles.card} style={{ marginBottom: 16 }}>
        <div className={styles.cardBody} style={{ textAlign: 'center', padding: '48px 24px' }}>
          <div style={{ width: 72, height: 72, borderRadius: '50%', background: '#DFF6DD', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 20px' }}>
            <Icon iconName="CheckMark" style={{ fontSize: 32, color: '#107C10' }} />
          </div>
          <Text variant="xLarge" style={{ color: '#107C10', fontWeight: 700, display: 'block', marginBottom: 10 }}>
            {submitSuccess}
          </Text>
          {savedId && (
            <Text variant="medium" style={{ color: '#605E5C', display: 'block', marginBottom: 24 }}>
              Dataverse ID: <code style={{ background: '#F3F2F1', padding: '2px 8px', borderRadius: 4, fontFamily: 'monospace', fontSize: 12 }}>{savedId}</code>
            </Text>
          )}
          <div style={{ display: 'flex', gap: 12, justifyContent: 'center', marginTop: 16 }}>
            <PrimaryButton
              text="Submit Another Request"
              iconProps={{ iconName: 'Add' }}
              onClick={() => this.setState({
                submitSuccess: null, savedId: null, errors: {},
                requestTitle: '', description: '', businessUnit: '',
                selectedSiteId: '', selectedSiteName: '', autoCostCentre: '', autoDivision: '', autoBudgetOwner: '',
                sctFunction: '', costIsFor: '', purchaseType: '',
                isNewSupplier: false, selectedSupplierId: '',
                newSupplierName: '', newSupplierLegalName: '', newSupplierContactName: '',
                newSupplierContactEmail: '', newSupplierContactPhone: '', newSupplierW9Notes: '',
                estimatedAmount: '', termStartDate: null, termEndDate: null, pricingBasis: '',
                dateNeededBy: null, urgency: '', urgencyReason: '', emergencyDeclaredBy: '',
                deliverTo: '', receivingContact: '', equipmentAssetTag: '',
                costRechargedToCustomer: false, customerName: '', customerAgreementAttached: false,
                techOrDataInvolved: false, supplierOnPremises: false, possibleCapitalPurchase: false,
                isConfidential: false, isRetrospective: false,
                declarationNoPriorCommitment: false, declarationNoConflictOfInterest: false,
                attachments: [], submitError: null,
                currentStep: 0,
              })}
              styles={{ root: { borderRadius: 6 } }}
            />
            {this.props.onFinished && (
              <DefaultButton
                text="View My Requests"
                iconProps={{ iconName: 'DocumentSearch' }}
                onClick={this.props.onFinished}
                styles={{ root: { borderRadius: 6 } }}
              />
            )}
          </div>
        </div>
      </div>
    );
  }

  // ── Main render ───────────────────────────────────────────────────────────────

  public render(): React.ReactElement {
    const { dataverseService } = this.props;
    const { loadingMeta, metaError, submitSuccess, submitting, savingDraft, submitError, uploadingFiles } = this.state;

    if (!dataverseService) return <LoadingSpinner label="Connecting to Dataverse..." />;
    if (loadingMeta) return <LoadingSpinner label="Loading sites and suppliers…" />;

    if (metaError) {
      return (
        <div className={styles.card} style={{ marginBottom: 16 }}>
          <div className={styles.cardBody}>
            <MessageBar messageBarType={MessageBarType.error} styles={{ root: { borderRadius: 6, marginBottom: 12 } }}>
              {metaError}
            </MessageBar>
            <DefaultButton text="Retry" iconProps={{ iconName: 'Refresh' }} onClick={() => void this._loadMeta()} styles={{ root: { borderRadius: 6 } }} />
          </div>
        </div>
      );
    }

    if (submitSuccess) return this._renderSuccess();

    const isWorking = submitting || savingDraft || uploadingFiles;

    // Hero header
    const hero = (
      <div className={styles.card} style={{
        marginBottom: 16,
        background: 'linear-gradient(135deg, #0F172A 0%, #1E293B 100%)',
        color: '#ffffff',
        border: 'none',
        boxShadow: '0 4px 16px rgba(15, 23, 42, 0.12)',
      }}>
        <div className={styles.cardBody} style={{ display: 'flex', alignItems: 'center', gap: 16, padding: '20px 24px' }}>
          <div style={{
            width: 48,
            height: 48,
            borderRadius: 12,
            background: 'linear-gradient(135deg, #0078D4, #2563EB)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            flexShrink: 0,
            boxShadow: '0 2px 8px rgba(0, 120, 212, 0.4)',
          }}>
            <Icon iconName="Add" style={{ fontSize: 24, color: '#ffffff' }} />
          </div>
          <div>
            <Text variant="xLarge" style={{ color: '#ffffff', fontWeight: 700, display: 'block', letterSpacing: '-0.2px' }}>
              New Spend Control Request
            </Text>
            <Text variant="medium" style={{ color: '#94A3B8', fontSize: 13 }}>
              Step {this.state.currentStep + 1} of 4: Complete required procurement details. Reference number is auto-generated upon submission.
            </Text>
          </div>
        </div>
      </div>
    );

    const { currentStep } = this.state;
    const STEPS = [
      { key: 0, title: 'Location & Requester', sub: 'Site, Unit & Line Manager' },
      { key: 1, title: 'Requirement & Spend', sub: 'Vendor, Amount & Type' },
      { key: 2, title: 'Governance & Logistics', sub: 'Asset, Recharges & Rules' },
      { key: 3, title: 'Review & Submit', sub: 'Evidence & Pre-Approvals' },
    ];
    const canNext = this._canGoNext(currentStep);

    // ── Step indicator ──
    const stepIndicator = (
      <div style={{
        background: '#ffffff',
        borderRadius: 12,
        border: '1px solid #E2E8F0',
        padding: '14px 20px',
        marginBottom: 20,
        boxShadow: '0 1px 4px rgba(15, 23, 42, 0.04)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        flexWrap: 'wrap',
        gap: 8,
      }}>
        {STEPS.map((sDef, idx) => {
          const isDone = idx < currentStep;
          const isActive = idx === currentStep;
          const isClickable = idx <= currentStep || (idx === currentStep + 1 && canNext);
          return (
            <React.Fragment key={idx}>
              <div
                onClick={() => {
                  if (isClickable) this.setState({ currentStep: idx });
                }}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: 10,
                  cursor: isClickable ? 'pointer' : 'default',
                  opacity: isActive || isDone ? 1 : 0.6,
                  padding: '6px 12px',
                  borderRadius: 8,
                  background: isActive ? '#EFF6FF' : 'transparent',
                  border: isActive ? '1px solid #BFDBFE' : '1px solid transparent',
                  transition: 'all 0.2s ease',
                }}
              >
                <div style={{
                  width: 30,
                  height: 30,
                  borderRadius: '50%',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  fontWeight: 700,
                  fontSize: 13,
                  background: isDone ? '#107C10' : isActive ? '#0078D4' : '#E2E8F0',
                  color: isDone || isActive ? '#ffffff' : '#64748B',
                  flexShrink: 0,
                  boxShadow: isActive ? '0 2px 6px rgba(0, 120, 212, 0.35)' : 'none',
                }}>
                  {isDone ? <Icon iconName="CheckMark" style={{ fontSize: 13 }} /> : idx + 1}
                </div>
                <div>
                  <div style={{
                    fontSize: 13,
                    fontWeight: isActive ? 700 : 600,
                    color: isActive ? '#0078D4' : isDone ? '#107C10' : '#334155',
                    lineHeight: 1.2,
                  }}>
                    {sDef.title}
                  </div>
                  <div style={{ fontSize: 11, color: '#64748B', marginTop: 2 }}>
                    {sDef.sub}
                  </div>
                </div>
              </div>
              {idx < STEPS.length - 1 && (
                <div style={{
                  flex: 1,
                  height: 2,
                  minWidth: 16,
                  background: idx < currentStep ? '#107C10' : '#E2E8F0',
                  borderRadius: 1,
                }} />
              )}
            </React.Fragment>
          );
        })}
      </div>
    );

    // ── Step content ──
    const stepContent = (() => {
      switch (currentStep) {
        case 0: return this._renderHeaderSection();
        case 1: return (
          <>
            {this._renderRequirementSection()}
            {this._renderFinancialSection()}
          </>
        );
        case 2: return (
          <>
            {this._renderLogisticsSection()}
            {this._renderDeclarationsSection()}
          </>
        );
        case 3: return (
          <>
            {this._renderAttachmentsSection()}
            {this._renderReviewStep()}
          </>
        );
        default: return null;
      }
    })();

    // ── Navigation footer ──
    const navFooter = (
      <div style={{ display: 'flex', gap: 12, padding: '16px 0 32px', alignItems: 'center' }}>
        {currentStep > 0 && (
          <DefaultButton
            text="Back"
            iconProps={{ iconName: 'ChevronLeft' }}
            disabled={isWorking}
            onClick={() => this.setState({ currentStep: currentStep - 1 })}
            styles={{ root: { borderRadius: 6, height: 40 } }}
          />
        )}

        {currentStep < STEPS.length - 1 && (
          <PrimaryButton
            text="Next →"
            iconProps={{ iconName: 'ChevronRight' }}
            disabled={!canNext || isWorking}
            onClick={() => this.setState({ currentStep: currentStep + 1 })}
            styles={{ root: { borderRadius: 6, height: 40, minWidth: 120 } }}
          />
        )}

        {currentStep === STEPS.length - 1 && (
          <>
            <PrimaryButton
              text={(submitting || this.state.isSubmitting) ? 'Submitting…' : 'Submit Request'}
              iconProps={{ iconName: (submitting || this.state.isSubmitting) ? undefined : 'Send' }}
              disabled={isWorking || this.state.isSubmitting || submitting || !this.state.declarationNoPriorCommitment || !this.state.declarationNoConflictOfInterest}
              title={(!this.state.declarationNoPriorCommitment || !this.state.declarationNoConflictOfInterest)
                ? 'Both mandatory declarations (R1 and R10) must be checked before submitting'
                : undefined}
              onClick={() => void this._submit(false)}
              onRenderText={(submitting || this.state.isSubmitting) ? () => (
                <span style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <Spinner size={SpinnerSize.xSmall} /> Submitting…
                </span>
              ) : undefined}
              styles={{ root: { borderRadius: 6, height: 40, minWidth: 160 } }}
            />
            <DefaultButton
              text={savingDraft ? 'Saving…' : 'Save as Draft'}
              iconProps={{ iconName: savingDraft ? undefined : 'Save' }}
              disabled={isWorking}
              onClick={() => void this._submit(true)}
              onRenderText={savingDraft ? () => (
                <span style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <Spinner size={SpinnerSize.xSmall} /> Saving…
                </span>
              ) : undefined}
              styles={{ root: { borderRadius: 6, height: 40 } }}
            />
          </>
        )}

        {/* Save Draft available on all steps */}
        {currentStep < STEPS.length - 1 && (
          <DefaultButton
            text={savingDraft ? 'Saving…' : 'Save Draft'}
            iconProps={{ iconName: 'Save' }}
            disabled={isWorking}
            onClick={() => void this._submit(true)}
            styles={{ root: { borderRadius: 6, height: 40 } }}
          />
        )}

        {uploadingFiles && (
          <span style={{ display: 'flex', alignItems: 'center', gap: 8, color: '#605E5C', fontSize: 13 }}>
            <Spinner size={SpinnerSize.xSmall} /> Uploading attachments…
          </span>
        )}
      </div>
    );

    return (
      <div>
        {hero}
        {stepIndicator}

        {submitError && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}
            onDismiss={() => this.setState({ submitError: null })}
            styles={{ root: { borderRadius: 6, marginBottom: 12 } }}>
            {submitError}
          </MessageBar>
        )}

        {stepContent}

        {/* Validation summary on Review step */}
        {currentStep === STEPS.length - 1 && Object.keys(this.state.errors).length > 0 && (
          <div style={{ background: '#FDE7E9', border: '1px solid #F4B8BC', borderRadius: 8, padding: '14px 18px', marginBottom: 16 }}>
            <Text variant="small" style={{ color: '#A4262C', fontWeight: 700, display: 'block', marginBottom: 8 }}>
              <Icon iconName="ErrorBadge" style={{ marginRight: 6, fontSize: 14 }} />
              Please fix the following before submitting:
            </Text>
            <ul style={{ margin: 0, paddingLeft: 20 }}>
              {Object.keys(this.state.errors).map((key, i) => (
                <li key={i} style={{ color: '#A4262C', fontSize: 13, marginBottom: 4 }}>{this.state.errors[key]}</li>
              ))}
            </ul>
          </div>
        )}

        {navFooter}
      </div>
    );
  }
}
