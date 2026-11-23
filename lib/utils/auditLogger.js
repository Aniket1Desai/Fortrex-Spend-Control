import { __awaiter, __generator } from "tslib";
/**
 * Human-readable labels for edit-panel fields used in audit descriptions.
 * Extend this map as more fields become editable.
 */
export var AUDIT_FIELD_LABELS = {
    sct_description: 'Description',
    sct_function: 'Function',
    sct_urgency_reason: 'Urgency Reason',
    sct_delivery_address: 'Delivery Address',
    sct_equipment_asset_tag: 'Equipment Asset Tag',
    sct_customer_name: 'Customer Name',
    sct_receiving_contact_email: 'Receiving Contact Email',
    sct_emergency_declared_by_email: 'Emergency Declared By',
    sct_estimated_total_value: 'Estimated Value',
    sct_date_needed_by: 'Date Needed By',
    sct_term_start_date: 'Term Start Date',
    sct_term_end_date: 'Term End Date',
    sct_cost_center: 'Cost Centre',
    sct_po_number: 'PO Number',
    sct_is_new_supplier: 'Is New Supplier',
    sct_cost_recharged_to_customer: 'Cost Recharged to Customer',
    sct_customer_agreement_attached: 'Customer Agreement Attached',
    sct_tech_or_data_involved: 'Tech or Data Involved',
    sct_supplier_on_premises: 'Supplier on Premises',
    sct_possible_capital_purchase: 'Possible Capital Purchase',
    sct_is_confidential: 'Confidential',
    sct_is_retrospective: 'Retrospective',
    sct_existing_agreement_or_po: 'Existing Agreement or PO',
    sct_declaration_no_prior_commitment: 'Declaration – No Prior Commitment',
    sct_declaration_no_conflict_of_interest: 'Declaration – No Conflict of Interest',
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
export function logAuditChanges(service, requestId, systemUserId, initialValues, currentValues) {
    return __awaiter(this, void 0, void 0, function () {
        var changedFields, tasks;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    changedFields = Object.keys(currentValues).filter(function (key) {
                        var _a, _b;
                        var before = (_a = initialValues[key]) !== null && _a !== void 0 ? _a : '';
                        var after = (_b = currentValues[key]) !== null && _b !== void 0 ? _b : '';
                        return before !== after;
                    });
                    if (changedFields.length === 0)
                        return [2 /*return*/];
                    tasks = changedFields.map(function (field) {
                        var _a, _b, _c;
                        var label = (_a = AUDIT_FIELD_LABELS[field]) !== null && _a !== void 0 ? _a : field;
                        var payload = {
                            'sct_requestid@odata.bind': "/sct_requests(".concat(requestId, ")"),
                            'sct_action': 2, // 2 = Edited
                            'sct_description': 'Updated field: ' + label,
                            'sct_oldvalue': (_b = initialValues[field]) !== null && _b !== void 0 ? _b : '',
                            'sct_newvalue': (_c = currentValues[field]) !== null && _c !== void 0 ? _c : '',
                            // No sct_timestamp — rely on Dataverse native createdon
                        };
                        if (systemUserId) {
                            payload['sct_actionby@odata.bind'] = "/systemusers(".concat(systemUserId, ")");
                        }
                        return service.createAuditLog(payload);
                    });
                    // Fire all POSTs concurrently; individual failures are non-fatal
                    return [4 /*yield*/, Promise.all(tasks.map(function (t, i) {
                            return t.catch(function (err) {
                                console.warn('[auditLogger] Failed to log change for field "' + changedFields[i] + '":', err);
                            });
                        }))];
                case 1:
                    // Fire all POSTs concurrently; individual failures are non-fatal
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
//# sourceMappingURL=auditLogger.js.map