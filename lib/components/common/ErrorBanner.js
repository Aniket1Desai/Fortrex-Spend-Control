import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { DefaultButton } from '@fluentui/react/lib/Button';
var ErrorBanner = function (_a) {
    var message = _a.message, onRetry = _a.onRetry, onDismiss = _a.onDismiss;
    return (React.createElement(MessageBar, { messageBarType: MessageBarType.error, isMultiline: false, onDismiss: onDismiss, dismissButtonAriaLabel: "Close", styles: {
            root: { borderRadius: 6, margin: '12px 0' },
        }, actions: onRetry ? (React.createElement("div", null,
            React.createElement(DefaultButton, { text: "Retry", onClick: onRetry, styles: { root: { height: 28, minWidth: 64 } } }))) : undefined }, message));
};
export default ErrorBanner;
//# sourceMappingURL=ErrorBanner.js.map