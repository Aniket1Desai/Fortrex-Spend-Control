import * as React from 'react';
import { DataverseService, ODataRecord } from '../../services/DataverseService';
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
    deleteTarget: {
        id: string;
        name: string;
    } | null;
}
export default class AttachmentManager extends React.Component<IAttachmentManagerProps, IAttachmentManagerState> {
    private _fileInput;
    constructor(props: IAttachmentManagerProps);
    componentDidMount(): Promise<void>;
    private _loadAttachments;
    private _onFileChange;
    private _deleteAttachment;
    private _confirmDelete;
    render(): React.ReactElement;
}
export {};
//# sourceMappingURL=AttachmentManager.d.ts.map