"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var Icon_1 = require("@fluentui/react/lib/Icon");
var Spinner_1 = require("@fluentui/react/lib/Spinner");
var MessageBar_1 = require("@fluentui/react/lib/MessageBar");
var Dialog_1 = require("@fluentui/react/lib/Dialog");
var Button_1 = require("@fluentui/react/lib/Button");
var DataverseService_1 = require("../../services/DataverseService");
var MAX_BYTES = 10 * 1024 * 1024; // 10 MB
function fileIcon(filename) {
    var _a, _b;
    var ext = (_b = (_a = filename.split('.').pop()) === null || _a === void 0 ? void 0 : _a.toLowerCase()) !== null && _b !== void 0 ? _b : '';
    if (ext === 'pdf')
        return 'PDF';
    if (ext === 'doc' || ext === 'docx')
        return 'WordDocument';
    if (ext === 'xls' || ext === 'xlsx')
        return 'ExcelDocument';
    if (['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'].indexOf(ext) !== -1)
        return 'Photo2';
    return 'Attach';
}
function formatSize(raw) {
    var n = parseInt(raw, 10);
    if (isNaN(n))
        return raw || '—';
    if (n >= 1024 * 1024)
        return (n / (1024 * 1024)).toFixed(1) + ' MB';
    if (n >= 1024)
        return Math.round(n / 1024) + ' KB';
    return n + ' B';
}
function formatDate(iso) {
    try {
        return new Intl.DateTimeFormat('en-AU', { day: '2-digit', month: 'short', year: 'numeric' }).format(new Date(iso));
    }
    catch (_a) {
        return iso;
    }
}
var AttachmentManager = /** @class */ (function (_super) {
    tslib_1.__extends(AttachmentManager, _super);
    function AttachmentManager(props) {
        var _this = _super.call(this, props) || this;
        _this._fileInput = React.createRef();
        _this._onFileChange = function (e) {
            var _a;
            var file = (_a = e.target.files) === null || _a === void 0 ? void 0 : _a[0];
            if (!file)
                return;
            // Reset input so same file can be re-uploaded
            e.target.value = '';
            if (file.size > MAX_BYTES) {
                _this.setState({ error: "File exceeds 10 MB limit (".concat((file.size / (1024 * 1024)).toFixed(1), " MB). Please compress or split the file.") });
                return;
            }
            var reader = new FileReader();
            reader.onload = function () {
                var result = reader.result;
                // Strip "data:mime/type;base64," prefix
                var base64 = result.indexOf(',') !== -1 ? result.split(',')[1] : result;
                _this.setState({ uploading: true, error: null });
                _this.props.dataverseService
                    .uploadAnnotation(_this.props.requestId, file.name, file.type, base64)
                    .then(function () { return _this._loadAttachments(); })
                    .then(function () { return _this.setState({ uploading: false }); })
                    .catch(function (err) {
                    _this.setState({ uploading: false, error: err instanceof Error ? err.message : String(err) });
                });
            };
            reader.onerror = function () {
                _this.setState({ error: 'Could not read the selected file.' });
            };
            reader.readAsDataURL(file);
        };
        _this._deleteAttachment = function (annotationId, filename) {
            _this.setState({ deleteTarget: { id: annotationId, name: filename } });
        };
        _this._confirmDelete = function () {
            var deleteTarget = _this.state.deleteTarget;
            if (!deleteTarget)
                return;
            _this.setState({ deleteTarget: null });
            _this.props.dataverseService
                .deleteAnnotation(deleteTarget.id)
                .then(function () { return _this._loadAttachments(); })
                .catch(function (err) {
                _this.setState({ error: err instanceof Error ? err.message : String(err) });
            });
        };
        _this.state = { attachments: [], loading: true, uploading: false, error: null, deleteTarget: null };
        return _this;
    }
    AttachmentManager.prototype.componentDidMount = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this._loadAttachments()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    AttachmentManager.prototype._loadAttachments = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var attachments, e_1;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.setState({ loading: true, error: null });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this.props.dataverseService.getAnnotations(this.props.requestId)];
                    case 2:
                        attachments = _a.sent();
                        this.setState({ attachments: attachments, loading: false });
                        return [3 /*break*/, 4];
                    case 3:
                        e_1 = _a.sent();
                        this.setState({ loading: false, error: e_1 instanceof Error ? e_1.message : String(e_1) });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    AttachmentManager.prototype.render = function () {
        var _this = this;
        var _a, _b;
        var readOnly = this.props.readOnly;
        var _c = this.state, attachments = _c.attachments, loading = _c.loading, uploading = _c.uploading, error = _c.error;
        var cardStyle = {
            border: '1px solid #EDEBE9', borderRadius: 8,
            padding: '14px 16px', marginBottom: 8,
            display: 'flex', alignItems: 'center', gap: 12,
            background: '#FAFAFA',
        };
        return (React.createElement("div", null,
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 10 } },
                React.createElement("div", { style: { fontSize: 12, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', letterSpacing: '0.5px' } },
                    "Attachments ",
                    !loading && "(".concat(attachments.length, ")")),
                !readOnly && (React.createElement(React.Fragment, null,
                    React.createElement("button", { onClick: function () { var _a; return (_a = _this._fileInput.current) === null || _a === void 0 ? void 0 : _a.click(); }, disabled: uploading, style: {
                            background: uploading ? '#EDEBE9' : '#0078D4',
                            color: uploading ? '#A19F9D' : '#fff',
                            border: 'none', borderRadius: 6,
                            padding: '5px 14px', fontSize: 12, fontWeight: 600,
                            cursor: uploading ? 'not-allowed' : 'pointer',
                            display: 'flex', alignItems: 'center', gap: 6,
                            fontFamily: 'inherit',
                        } }, uploading
                        ? React.createElement(React.Fragment, null,
                            React.createElement(Spinner_1.Spinner, { size: Spinner_1.SpinnerSize.xSmall }),
                            " Uploading\u2026")
                        : React.createElement(React.Fragment, null,
                            React.createElement(Icon_1.Icon, { iconName: "Attach", style: { fontSize: 13 } }),
                            " Attach File")),
                    React.createElement("input", { ref: this._fileInput, type: "file", accept: ".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg,.gif,.bmp,.webp", style: { display: 'none' }, onChange: this._onFileChange })))),
            error && (React.createElement(MessageBar_1.MessageBar, { messageBarType: MessageBar_1.MessageBarType.error, onDismiss: function () { return _this.setState({ error: null }); }, styles: { root: { borderRadius: 6, marginBottom: 10 } } }, error)),
            loading && (React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 8, padding: '12px 0', color: '#605E5C', fontSize: 13 } },
                React.createElement(Spinner_1.Spinner, { size: Spinner_1.SpinnerSize.small }),
                " Loading attachments\u2026")),
            !loading && attachments.length === 0 && (React.createElement("div", { style: { textAlign: 'center', padding: '20px 0', color: '#A19F9D', fontSize: 13 } },
                React.createElement(Icon_1.Icon, { iconName: "Attach", style: { fontSize: 28, display: 'block', margin: '0 auto 6px', color: '#C8C6C4' } }),
                "No attachments yet",
                readOnly ? '.' : ' — click Attach File to upload.')),
            !loading && attachments.map(function (a) {
                var _a;
                var annotationId = String((_a = a['annotationid']) !== null && _a !== void 0 ? _a : '');
                var filename = (0, DataverseService_1.strVal)(a, 'filename') || 'Unnamed file';
                var filesize = formatSize((0, DataverseService_1.strVal)(a, 'filesize'));
                var createdOn = formatDate((0, DataverseService_1.strVal)(a, 'createdon'));
                var icon = fileIcon(filename);
                return (React.createElement("div", { key: annotationId, style: cardStyle },
                    React.createElement("div", { style: {
                            width: 36, height: 36, borderRadius: 8, background: '#DEECF9',
                            display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0,
                        } },
                        React.createElement(Icon_1.Icon, { iconName: icon, style: { fontSize: 18, color: '#0078D4' } })),
                    React.createElement("div", { style: { flex: 1, minWidth: 0 } },
                        React.createElement("div", { style: { fontSize: 13, fontWeight: 600, color: '#201F1E', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, filename),
                        React.createElement("div", { style: { fontSize: 11, color: '#A19F9D', marginTop: 2 } },
                            filesize,
                            " \u00B7 Uploaded ",
                            createdOn)),
                    React.createElement("div", { style: { display: 'flex', gap: 6, flexShrink: 0 } },
                        React.createElement("button", { disabled: true, title: "Download via SharePoint record view", style: {
                                background: '#F3F2F1', color: '#A19F9D', border: 'none',
                                borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600,
                                cursor: 'not-allowed', fontFamily: 'inherit',
                                display: 'flex', alignItems: 'center', gap: 4,
                            } },
                            React.createElement(Icon_1.Icon, { iconName: "View", style: { fontSize: 11 } }),
                            " Preview"),
                        !readOnly && (React.createElement("button", { onClick: function () { return _this._deleteAttachment(annotationId, filename); }, title: "Remove attachment", style: {
                                background: '#FDE7E9', color: '#A4262C', border: 'none',
                                borderRadius: 6, padding: '4px 8px', fontSize: 11,
                                cursor: 'pointer', fontFamily: 'inherit',
                                display: 'flex', alignItems: 'center',
                            } },
                            React.createElement(Icon_1.Icon, { iconName: "Delete", style: { fontSize: 12 } }))))));
            }),
            !readOnly && (React.createElement("div", { style: { fontSize: 11, color: '#C8C6C4', marginTop: 6 } }, "Supported: PDF, Word, Excel, images \u00B7 Max 10 MB per file")),
            React.createElement(Dialog_1.Dialog, { hidden: !this.state.deleteTarget, onDismiss: function () { return _this.setState({ deleteTarget: null }); }, dialogContentProps: {
                    type: Dialog_1.DialogType.normal,
                    title: 'Remove Attachment',
                    subText: "Remove \"".concat((_b = (_a = this.state.deleteTarget) === null || _a === void 0 ? void 0 : _a.name) !== null && _b !== void 0 ? _b : '', "\"? This action cannot be undone."),
                }, modalProps: { isBlocking: true }, minWidth: 440 },
                React.createElement(Dialog_1.DialogFooter, null,
                    React.createElement(Button_1.PrimaryButton, { text: "Remove", iconProps: { iconName: 'Delete' }, onClick: this._confirmDelete, styles: { root: { background: '#A4262C', borderColor: '#A4262C', borderRadius: 6 } } }),
                    React.createElement(Button_1.DefaultButton, { text: "Cancel", onClick: function () { return _this.setState({ deleteTarget: null }); }, styles: { root: { borderRadius: 6 } } })))));
    };
    return AttachmentManager;
}(React.Component));
exports.default = AttachmentManager;
//# sourceMappingURL=AttachmentManager.js.map