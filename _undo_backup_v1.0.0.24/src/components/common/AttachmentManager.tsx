import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { DataverseService, ODataRecord, strVal } from '../../services/DataverseService';

interface IAttachmentManagerProps {
  dataverseService: DataverseService;
  requestId: string;
  readOnly?: boolean;
}

interface IAttachmentManagerState {
  attachments: ODataRecord[];
  loading: boolean;
  uploading: boolean;
  error: string | null;
  deleteTarget: { id: string; name: string } | null;
}

const MAX_BYTES = 10 * 1024 * 1024; // 10 MB

function fileIcon(filename: string): string {
  const ext = filename.split('.').pop()?.toLowerCase() ?? '';
  if (ext === 'pdf') return 'PDF';
  if (ext === 'doc' || ext === 'docx') return 'WordDocument';
  if (ext === 'xls' || ext === 'xlsx') return 'ExcelDocument';
  if (['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'].indexOf(ext) !== -1) return 'Photo2';
  return 'Attach';
}

function formatSize(raw: string): string {
  const n = parseInt(raw, 10);
  if (isNaN(n)) return raw || '—';
  if (n >= 1024 * 1024) return (n / (1024 * 1024)).toFixed(1) + ' MB';
  if (n >= 1024) return Math.round(n / 1024) + ' KB';
  return n + ' B';
}

function formatDate(iso: string): string {
  try {
    return new Intl.DateTimeFormat('en-AU', { day: '2-digit', month: 'short', year: 'numeric' }).format(new Date(iso));
  } catch { return iso; }
}

export default class AttachmentManager extends React.Component<IAttachmentManagerProps, IAttachmentManagerState> {
  private _fileInput: React.RefObject<HTMLInputElement> = React.createRef();

  constructor(props: IAttachmentManagerProps) {
    super(props);
    this.state = { attachments: [], loading: true, uploading: false, error: null, deleteTarget: null };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadAttachments();
  }

  private async _loadAttachments(): Promise<void> {
    this.setState({ loading: true, error: null });
    try {
      const attachments = await this.props.dataverseService.getAnnotations(this.props.requestId);
      this.setState({ attachments, loading: false });
    } catch (e: unknown) {
      this.setState({ loading: false, error: e instanceof Error ? e.message : String(e) });
    }
  }

  private _onFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (!file) return;
    // Reset input so same file can be re-uploaded
    e.target.value = '';

    if (file.size > MAX_BYTES) {
      this.setState({ error: `File exceeds 10 MB limit (${(file.size / (1024 * 1024)).toFixed(1)} MB). Please compress or split the file.` });
      return;
    }

    const reader = new FileReader();
    reader.onload = (): void => {
      const result = reader.result as string;
      // Strip "data:mime/type;base64," prefix
      const base64 = result.indexOf(',') !== -1 ? result.split(',')[1] : result;
      this.setState({ uploading: true, error: null });
      this.props.dataverseService
        .uploadAnnotation(this.props.requestId, file.name, file.type, base64)
        .then(() => this._loadAttachments())
        .then(() => this.setState({ uploading: false }))
        .catch((err: unknown) => {
          this.setState({ uploading: false, error: err instanceof Error ? err.message : String(err) });
        });
    };
    reader.onerror = (): void => {
      this.setState({ error: 'Could not read the selected file.' });
    };
    reader.readAsDataURL(file);
  };

  private _deleteAttachment = (annotationId: string, filename: string): void => {
    this.setState({ deleteTarget: { id: annotationId, name: filename } });
  };

  private _confirmDelete = (): void => {
    const { deleteTarget } = this.state;
    if (!deleteTarget) return;
    this.setState({ deleteTarget: null });
    this.props.dataverseService
      .deleteAnnotation(deleteTarget.id)
      .then(() => this._loadAttachments())
      .catch((err: unknown) => {
        this.setState({ error: err instanceof Error ? err.message : String(err) });
      });
  };

  public render(): React.ReactElement {
    const { readOnly } = this.props;
    const { attachments, loading, uploading, error } = this.state;

    const cardStyle: React.CSSProperties = {
      border: '1px solid #EDEBE9', borderRadius: 8,
      padding: '14px 16px', marginBottom: 8,
      display: 'flex', alignItems: 'center', gap: 12,
      background: '#FAFAFA',
    };

    return (
      <div>
        {/* Header row */}
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 10 }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: '#A19F9D', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
            Attachments {!loading && `(${attachments.length})`}
          </div>
          {!readOnly && (
            <>
              <button
                onClick={() => this._fileInput.current?.click()}
                disabled={uploading}
                style={{
                  background: uploading ? '#EDEBE9' : '#0078D4',
                  color: uploading ? '#A19F9D' : '#fff',
                  border: 'none', borderRadius: 6,
                  padding: '5px 14px', fontSize: 12, fontWeight: 600,
                  cursor: uploading ? 'not-allowed' : 'pointer',
                  display: 'flex', alignItems: 'center', gap: 6,
                  fontFamily: 'inherit',
                }}
              >
                {uploading
                  ? <><Spinner size={SpinnerSize.xSmall} /> Uploading…</>
                  : <><Icon iconName="Attach" style={{ fontSize: 13 }} /> Attach File</>
                }
              </button>
              <input
                ref={this._fileInput}
                type="file"
                accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg,.gif,.bmp,.webp"
                style={{ display: 'none' }}
                onChange={this._onFileChange}
              />
            </>
          )}
        </div>

        {/* Error */}
        {error && (
          <MessageBar messageBarType={MessageBarType.error}
            onDismiss={() => this.setState({ error: null })}
            styles={{ root: { borderRadius: 6, marginBottom: 10 } }}>
            {error}
          </MessageBar>
        )}

        {/* Loading */}
        {loading && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '12px 0', color: '#605E5C', fontSize: 13 }}>
            <Spinner size={SpinnerSize.small} /> Loading attachments…
          </div>
        )}

        {/* Empty state */}
        {!loading && attachments.length === 0 && (
          <div style={{ textAlign: 'center', padding: '20px 0', color: '#A19F9D', fontSize: 13 }}>
            <Icon iconName="Attach" style={{ fontSize: 28, display: 'block', margin: '0 auto 6px', color: '#C8C6C4' }} />
            No attachments yet{readOnly ? '.' : ' — click Attach File to upload.'}
          </div>
        )}

        {/* Attachment list */}
        {!loading && attachments.map(a => {
          const annotationId = String(a['annotationid'] ?? '');
          const filename     = strVal(a, 'filename') || 'Unnamed file';
          const filesize     = formatSize(strVal(a, 'filesize'));
          const createdOn    = formatDate(strVal(a, 'createdon'));
          const icon         = fileIcon(filename);

          return (
            <div key={annotationId} style={cardStyle}>
              {/* File icon */}
              <div style={{
                width: 36, height: 36, borderRadius: 8, background: '#DEECF9',
                display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0,
              }}>
                <Icon iconName={icon} style={{ fontSize: 18, color: '#0078D4' }} />
              </div>

              {/* File details */}
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: '#201F1E', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {filename}
                </div>
                <div style={{ fontSize: 11, color: '#A19F9D', marginTop: 2 }}>
                  {filesize} · Uploaded {createdOn}
                </div>
              </div>

              {/* Actions */}
              <div style={{ display: 'flex', gap: 6, flexShrink: 0 }}>
                <button
                  disabled
                  title="Download via SharePoint record view"
                  style={{
                    background: '#F3F2F1', color: '#A19F9D', border: 'none',
                    borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 600,
                    cursor: 'not-allowed', fontFamily: 'inherit',
                    display: 'flex', alignItems: 'center', gap: 4,
                  }}
                >
                  <Icon iconName="View" style={{ fontSize: 11 }} /> Preview
                </button>
                {!readOnly && (
                  <button
                    onClick={() => this._deleteAttachment(annotationId, filename)}
                    title="Remove attachment"
                    style={{
                      background: '#FDE7E9', color: '#A4262C', border: 'none',
                      borderRadius: 6, padding: '4px 8px', fontSize: 11,
                      cursor: 'pointer', fontFamily: 'inherit',
                      display: 'flex', alignItems: 'center',
                    }}
                  >
                    <Icon iconName="Delete" style={{ fontSize: 12 }} />
                  </button>
                )}
              </div>
            </div>
          );
        })}

        {!readOnly && (
          <div style={{ fontSize: 11, color: '#C8C6C4', marginTop: 6 }}>
            Supported: PDF, Word, Excel, images · Max 10 MB per file
          </div>
        )}

        {/* Fluent UI delete confirmation dialog */}
        <Dialog
          hidden={!this.state.deleteTarget}
          onDismiss={() => this.setState({ deleteTarget: null })}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Remove Attachment',
            subText: `Remove "${this.state.deleteTarget?.name ?? ''}"? This action cannot be undone.`,
          }}
          modalProps={{ isBlocking: true }}
          minWidth={440}
        >
          <DialogFooter>
            <PrimaryButton
              text="Remove"
              iconProps={{ iconName: 'Delete' }}
              onClick={this._confirmDelete}
              styles={{ root: { background: '#A4262C', borderColor: '#A4262C', borderRadius: 6 } }}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ deleteTarget: null })}
              styles={{ root: { borderRadius: 6 } }}
            />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
