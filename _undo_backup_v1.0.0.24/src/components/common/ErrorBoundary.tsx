import * as React from 'react';

interface IProps { children?: React.ReactNode; }
interface IState { error: Error | null; }

export default class ErrorBoundary extends React.Component<IProps, IState> {
  constructor(props: IProps) {
    super(props);
    this.state = { error: null };
  }

  public static getDerivedStateFromError(error: Error): IState {
    return { error };
  }

  public render(): React.ReactNode {
    const { error } = this.state;
    if (!error) return this.props.children;
    return (
      <div style={{
        margin: 24, padding: 24, background: '#FFF4CE',
        border: '1px solid #F7C948', borderRadius: 8, fontFamily: 'Segoe UI, sans-serif',
      }}>
        <strong style={{ color: '#6B3A00', fontSize: 15 }}>Something went wrong</strong>
        <p style={{ color: '#6B3A00', marginTop: 8, fontSize: 13 }}>{error.message}</p>
        <button
          style={{ marginTop: 12, padding: '6px 16px', borderRadius: 6, border: '1px solid #C7A333', background: '#fff', cursor: 'pointer', fontSize: 13 }}
          onClick={() => this.setState({ error: null })}
        >
          Try again
        </button>
      </div>
    );
  }
}
