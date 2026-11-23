"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var Panel_1 = require("@fluentui/react/lib/Panel");
var Icon_1 = require("@fluentui/react/lib/Icon");
var LogsContent_1 = tslib_1.__importDefault(require("./LogsContent"));
var RequestLogsPanel = function (_a) {
    var requestId = _a.requestId, reference = _a.reference, service = _a.service, onDismiss = _a.onDismiss;
    return (React.createElement(Panel_1.Panel, { isOpen: true, onDismiss: onDismiss, type: Panel_1.PanelType.custom, customWidth: "min(760px, 96vw)", hasCloseButton: true, closeButtonAriaLabel: "Close history panel", onRenderHeader: function () { return (React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 12, padding: '18px 24px 0', flex: 1, minWidth: 0 } },
            React.createElement("div", { style: {
                    width: 36, height: 36, borderRadius: 8, flexShrink: 0,
                    background: 'linear-gradient(135deg, #107C10 0%, #8764B8 100%)',
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                } },
                React.createElement(Icon_1.Icon, { iconName: "History", style: { fontSize: 16, color: '#fff' } })),
            React.createElement("div", { style: { minWidth: 0 } },
                React.createElement("div", { style: { fontSize: 16, fontWeight: 700, color: '#201F1E', lineHeight: '1.2' } }, "Request History"),
                reference && (React.createElement("div", { style: { fontSize: 12, color: '#605E5C', marginTop: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, reference))))); }, styles: {
            main: { boxShadow: '-4px 0 20px rgba(0,0,0,0.10)' },
            scrollableContent: { display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' },
            content: { flex: 1, overflowY: 'auto', padding: '16px 24px 32px' },
        } },
        React.createElement(LogsContent_1.default, { requestId: requestId, service: service })));
};
exports.default = RequestLogsPanel;
//# sourceMappingURL=RequestLogsPanel.js.map