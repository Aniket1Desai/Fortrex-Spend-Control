import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { DefaultButton } from '@fluentui/react/lib/Button';

interface IEmptyStateProps {
  message: string;
  subMessage?: string;
  iconName?: string;
  actionLabel?: string;
  onAction?: () => void;
}

const EmptyState: React.FC<IEmptyStateProps> = ({
  message,
  subMessage,
  iconName = 'DocumentSet',
  actionLabel,
  onAction,
}) => (
  <div
    style={{
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      padding: '56px 24px',
      minHeight: 240,
      textAlign: 'center',
    }}
  >
    <div
      style={{
        width: 72,
        height: 72,
        borderRadius: '50%',
        background: 'linear-gradient(135deg, #EBF3FB 0%, #D0E8F8 100%)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        marginBottom: 20,
      }}
    >
      <Icon
        iconName={iconName}
        style={{ fontSize: 32, color: '#0078D4' }}
      />
    </div>
    <Text
      variant="xLarge"
      style={{ color: '#323130', fontWeight: 600, marginBottom: 8 }}
    >
      {message}
    </Text>
    {subMessage && (
      <Text variant="medium" style={{ color: '#605E5C', maxWidth: 380, lineHeight: '1.5' }}>
        {subMessage}
      </Text>
    )}
    {actionLabel && onAction && (
      <DefaultButton
        text={actionLabel}
        onClick={onAction}
        style={{ marginTop: 24 }}
        styles={{
          root: { borderColor: '#0078D4', color: '#0078D4' },
          rootHovered: { background: '#EBF3FB' },
        }}
      />
    )}
  </div>
);

export default EmptyState;
