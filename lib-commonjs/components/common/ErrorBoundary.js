"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var ErrorBoundary = /** @class */ (function (_super) {
    tslib_1.__extends(ErrorBoundary, _super);
    function ErrorBoundary(props) {
        var _this = _super.call(this, props) || this;
        _this.state = { error: null };
        return _this;
    }
    ErrorBoundary.getDerivedStateFromError = function (error) {
        return { error: error };
    };
    ErrorBoundary.prototype.render = function () {
        var _this = this;
        var error = this.state.error;
        if (!error)
            return this.props.children;
        return (React.createElement("div", { style: {
                margin: 24, padding: 24, background: '#FFF4CE',
                border: '1px solid #F7C948', borderRadius: 8, fontFamily: 'Segoe UI, sans-serif',
            } },
            React.createElement("strong", { style: { color: '#6B3A00', fontSize: 15 } }, "Something went wrong"),
            React.createElement("p", { style: { color: '#6B3A00', marginTop: 8, fontSize: 13 } }, error.message),
            React.createElement("button", { style: { marginTop: 12, padding: '6px 16px', borderRadius: 6, border: '1px solid #C7A333', background: '#fff', cursor: 'pointer', fontSize: 13 }, onClick: function () { return _this.setState({ error: null }); } }, "Try again")));
    };
    return ErrorBoundary;
}(React.Component));
exports.default = ErrorBoundary;
//# sourceMappingURL=ErrorBoundary.js.map