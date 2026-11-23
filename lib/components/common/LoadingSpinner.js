import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
var LoadingSpinner = function (_a) {
    var _b = _a.label, label = _b === void 0 ? 'Loading...' : _b, _c = _a.size, size = _c === void 0 ? SpinnerSize.large : _c;
    return (React.createElement("div", { style: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '48px 24px',
            minHeight: 200,
        } },
        React.createElement(Spinner, { size: size, label: label, labelPosition: "bottom" })));
};
export default LoadingSpinner;
//# sourceMappingURL=LoadingSpinner.js.map