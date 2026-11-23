import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { DefaultButton } from '@fluentui/react/lib/Button';
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
            React.createElement(Icon, { iconName: iconName, style: { fontSize: 32, color: '#0078D4' } })),
        React.createElement(Text, { variant: "xLarge", style: { color: '#323130', fontWeight: 600, marginBottom: 8 } }, message),
        subMessage && (React.createElement(Text, { variant: "medium", style: { color: '#605E5C', maxWidth: 380, lineHeight: '1.5' } }, subMessage)),
        actionLabel && onAction && (React.createElement(DefaultButton, { text: actionLabel, onClick: onAction, style: { marginTop: 24 }, styles: {
                root: { borderColor: '#0078D4', color: '#0078D4' },
                rootHovered: { background: '#EBF3FB' },
            } }))));
};
export default EmptyState;
//# sourceMappingURL=EmptyState.js.map