"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var Spinner_1 = require("@fluentui/react/lib/Spinner");
var LoadingSpinner = function (_a) {
    var _b = _a.label, label = _b === void 0 ? 'Loading...' : _b, _c = _a.size, size = _c === void 0 ? Spinner_1.SpinnerSize.large : _c;
    return (React.createElement("div", { style: {
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '48px 24px',
            minHeight: 200,
        } },
        React.createElement(Spinner_1.Spinner, { size: size, label: label, labelPosition: "bottom" })));
};
exports.default = LoadingSpinner;
//# sourceMappingURL=LoadingSpinner.js.map