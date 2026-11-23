"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var Icon_1 = require("@fluentui/react/lib/Icon");
var Text_1 = require("@fluentui/react/lib/Text");
var Button_1 = require("@fluentui/react/lib/Button");
var EmptyState = function (_a) {
    var message = _a.message, subMessage = _a.subMessage, _b = _a.iconName, iconName = _b === void 0 ? 'DocumentSet' : _b, actionLabel = _a.actionLabel, onAction = _a.onAction;
    return (React.createElement("div", { style: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '56px 24px',
            minHeight: 240,
            textAlign: 'center',
        } },
        React.createElement("div", { style: {
                width: 72,
                height: 72,
                borderRadius: '50%',
                background: 'linear-gradient(135deg, #EBF3FB 0%, #D0E8F8 100%)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                marginBottom: 20,
            } },
            React.createElement(Icon_1.Icon, { iconName: iconName, style: { fontSize: 32, color: '#0078D4' } })),
        React.createElement(Text_1.Text, { variant: "xLarge", style: { color: '#323130', fontWeight: 600, marginBottom: 8 } }, message),
        subMessage && (React.createElement(Text_1.Text, { variant: "medium", style: { color: '#605E5C', maxWidth: 380, lineHeight: '1.5' } }, subMessage)),
        actionLabel && onAction && (React.createElement(Button_1.DefaultButton, { text: actionLabel, onClick: onAction, style: { marginTop: 24 }, styles: {
                root: { borderColor: '#0078D4', color: '#0078D4' },
                rootHovered: { background: '#EBF3FB' },
            } }))));
};
exports.default = EmptyState;
//# sourceMappingURL=EmptyState.js.map