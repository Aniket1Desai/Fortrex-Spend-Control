import { DataverseService } from '../services/DataverseService';

/**
 * Human-readable labels for edit-panel fields used in audit descriptions.
 * Extend this map as more fields become editable.
 */
export const AUDIT_FIELD_LABELS: Record<string, string> = {
  sct_description:                   'Description',
  sct_function:                      'Function',
  sct_urgency_reason:                'Urgency Reason',
  sct_delivery_address:              'Delivery Address',
  sct_equipment_asset_tag:           'Equipment Asset Tag',
  sct_customer_name:                 'Customer Name',
  sct_receiving_contact_email:       'Receiving Contact Email',
  sct_emergency_declared_by_email:   'Emergency Declared By',
  sct_estimated_total_value:         'Estimated Value',
  sct_date_needed_by:                'Date Needed By',
  sct_term_start_date:               'Term Start Date',
  sct_term_end_date:                 'Term End Date',
  sct_cost_center:                   'Cost Centre',
  sct_po_number:                     'PO Number',
  sct_is_new_supplier:               'Is New Supplier',
  sct_cost_recharged_to_customer:    'Cost Recharged to Customer',
  sct_customer_agreement_attached:   'Customer Agreement Attached',
  sct_tech_or_data_involved:         'Tech or Data Involved',
  sct_supplier_on_premises:          'Supplier on Premises',
  sct_possible_capital_purchase:     'Possible Capital Purchase',
  sct_is_confidential:               'Confidential',
  sct_is_retrospective:              'Retrospective',
  sct_existing_agreement_or_po:      'Existing Agreement or PO',
  sct_declaration_no_prior_commitment:    'Declaration – No Prior Commitment',
  sct_declaration_no_conflict_of_interest:'Declaration – No Conflict of Interest',
};

/**
 * Compares initialValues with currentValues and POSTs one sct_auditlog record per
 * changed field, in parallel via Promise.all.
 *
 * @param service       DataverseService instance
 * @param requestId     GUID of the sct_request being edited
 * @param systemUserId  GUID of the current systemuser (null = omit the @odata.bind lookup)
 * @param initialValues Field values as loaded when the edit panel was opened
 * @param currentValues Field values as they were when Save was clicked
 */
export async function logAuditChanges(
  service: DataverseService,
  requestId: string,
  systemUserId: string | null,
  initialValues: Record<string, string>,
  currentValues: Record<string, string>
): Promise<void> {
  const changedFields = Object.keys(currentValues).filter(key => {
    const before = initialValues[key] ?? '';
    const after  = currentValues[key]  ?? '';
    return before !== after;
  });

  if (changedFields.length === 0) return;

  const tasks = changedFields.map(field => {
    const label = AUDIT_FIELD_LABELS[field] ?? field;
    const payload: Record<string, unknown> = {
      'sct_requestid@odata.bind': `/sct_requests(${requestId})`,
      'sct_action': 2,                                    // 2 = Edited
      'sct_description': 'Updated field: ' + label,
      'sct_oldvalue':    initialValues[field] ?? '',
      'sct_newvalue':    currentValues[field]  ?? '',
      // No sct_timestamp — rely on Dataverse native createdon
    };
    if (systemUserId) {
      payload['sct_actionby@odata.bind'] = `/systemusers(${systemUserId})`;
    }
    return service.createAuditLog(payload);
  });

  // Fire all POSTs concurrently; individual failures are non-fatal
  await Promise.all(
    tasks.map((t, i) =>
      t.catch((err: unknown) => {
        console.warn('[auditLogger] Failed to log change for field "' + changedFields[i] + '":', err);
      })
    )
  );
}
