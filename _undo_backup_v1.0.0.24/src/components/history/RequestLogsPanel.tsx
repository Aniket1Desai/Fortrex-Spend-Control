import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Icon } from '@fluentui/react/lib/Icon';

import { DataverseService } from '../../services/DataverseService';
import LogsContent from './LogsContent';

interface IRequestLogsPanelProps {
  requestId: string;
  reference: string;
  service: DataverseService;
  onDismiss: () => void;
}

const RequestLogsPanel: React.FC<IRequestLogsPanelProps> = ({
  requestId, reference, service, onDismiss,
}) => (
  <Panel
    isOpen
    onDismiss={onDismiss}
    type={PanelType.custom}
    customWidth="min(760px, 96vw)"
    hasCloseButton
    closeButtonAriaLabel="Close history panel"
    onRenderHeader={() => (
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '18px 24px 0', flex: 1, minWidth: 0 }}>
        <div style={{
          width: 36, height: 36, borderRadius: 8, flexShrink: 0,
          background: 'linear-gradient(135deg, #107C10 0%, #8764B8 100%)',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}>
          <Icon iconName="History" style={{ fontSize: 16, color: '#fff' }} />
        </div>
        <div style={{ minWidth: 0 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: '#201F1E', lineHeight: '1.2' }}>
            Request History
          </div>
          {reference && (
            <div style={{ fontSize: 12, color: '#605E5C', marginTop: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {reference}
            </div>
          )}
        </div>
      </div>
    )}
    styles={{
      main: { boxShadow: '-4px 0 20px rgba(0,0,0,0.10)' },
      scrollableContent: { display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' },
      content: { flex: 1, overflowY: 'auto', padding: '16px 24px 32px' },
    }}
  >
    <LogsContent requestId={requestId} service={service} />
  </Panel>
);

export default RequestLogsPanel;
