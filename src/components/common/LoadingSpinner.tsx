import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

interface ILoadingSpinnerProps {
  label?: string;
  size?: SpinnerSize;
}

const LoadingSpinner: React.FC<ILoadingSpinnerProps> = ({
  label = 'Loading...',
  size = SpinnerSize.large,
}) => (
  <div
    style={{
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      padding: '48px 24px',
      minHeight: 200,
    }}
  >
    <Spinner size={size} label={label} labelPosition="bottom" />
  </div>
);

export default LoadingSpinner;
