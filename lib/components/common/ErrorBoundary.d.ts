import * as React from 'react';
interface IProps {
    children?: React.ReactNode;
}
interface IState {
    error: Error | null;
}
export default class ErrorBoundary extends React.Component<IProps, IState> {
    constructor(props: IProps);
    static getDerivedStateFromError(error: Error): IState;
    render(): React.ReactNode;
}
export {};
//# sourceMappingURL=ErrorBoundary.d.ts.map