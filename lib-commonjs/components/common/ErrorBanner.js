"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var MessageBar_1 = require("@fluentui/react/lib/MessageBar");
var Button_1 = require("@fluentui/react/lib/Button");
var ErrorBanner = function (_a) {
    var message = _a.message, onRetry = _a.onRetry, onDismiss = _a.onDismiss;
    return (React.createElement(MessageBar_1.MessageBar, { messageBarType: MessageBar_1.MessageBarType.error, isMultiline: false, onDismiss: onDismiss, dismissButtonAriaLabel: "Close", styles: {
            root: { borderRadius: 6, margin: '12px 0' },
        }, actions: onRetry ? (React.createElement("div", null,
            React.createElement(Button_1.DefaultButton, { text: "Retry", onClick: onRetry, styles: { root: { height: 28, minWidth: 64 } } }))) : undefined }, message));
};
exports.default = ErrorBanner;
//# sourceMappingURL=ErrorBanner.js.map