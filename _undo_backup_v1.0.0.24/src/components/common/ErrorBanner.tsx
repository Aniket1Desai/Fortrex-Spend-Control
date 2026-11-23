import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { DefaultButton } from '@fluentui/react/lib/Button';

interface IErrorBannerProps {
  message: string;
  onRetry?: () => void;
  onDismiss?: () => void;
}

const ErrorBanner: React.FC<IErrorBannerProps> = ({ message, onRetry, onDismiss }) => (
  <MessageBar
    messageBarType={MessageBarType.error}
    isMultiline={false}
    onDismiss={onDismiss}
    dismissButtonAriaLabel="Close"
    styles={{
      root: { borderRadius: 6, margin: '12px 0' },
    }}
    actions={
      onRetry ? (
        <div>
          <DefaultButton
            text="Retry"
            onClick={onRetry}
            styles={{ root: { height: 28, minWidth: 64 } }}
          />
        </div>
      ) : undefined
    }
  >
    {message}
  </MessageBar>
);

export default ErrorBanner;
