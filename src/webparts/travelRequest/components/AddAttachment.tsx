import * as React from 'react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import styles from './TravelRequest.module.scss';
import { DataService } from '../../../services/data-service';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { PrimaryButton, ActionButton } from 'office-ui-fabric-react/lib/Button';
import { stringIsNullOrEmpty } from '@pnp/common';


export interface IAddAttachmentProps {
    isOpen: boolean;
    context: WebPartContext;
    onClose(success?: boolean): void;
    formKey: string;
}

export interface IAddAttachmentState {
    formError: boolean;
    formMessage: string;
    files: File[];
    uploading: boolean;
    photoLibUrl: string;
}


export default class AddAttachment extends React.Component<IAddAttachmentProps, IAddAttachmentState>{
    private service: DataService;
    constructor(props) {
        super(props);

        this.state = {
            formError: false,
            formMessage: "",
            files: [],
            uploading: false,
            photoLibUrl: "",
        };

        this.service = new DataService(this.props.context.pageContext);
    }

    /* public componentDidMount() {
     }
 
     public componentDidUpdate(prevProps: IAddAttachmentProps) {
     
     }
 */



    //file input
    private handleFileChange(event) {
        const { name, value } = event.target;
        var partialState = {};
        partialState[name] = event.target.files;
        this.setState(partialState);
    }



    //submit files
    private async postFiles() {
        this.setState({ uploading: true });
        try {
            let result = await this.service.AddAttachments(this.state, this.props.formKey);
            this._onClose(true);
        }
        catch (e) {
            this.setState({
                formError: true,
                formMessage: "We are having trouble adding your attachment.",
                uploading: false
            });
        }
    }


    private _onClose(success?: boolean) {
        this.setState({
            files: [],
            formError: false,
            formMessage: "",
            uploading: false,
        });
        success ? this.props.onClose(success) : this.props.onClose();
    }

    public render(): React.ReactElement<IAddAttachmentProps> {
        const { isOpen , formKey} = this.props;
        const { formError, formMessage,  files, uploading, } = this.state;
        const disabled = () => {
            if (files.length == 0) {
                return true;
            }
            else {
                return false;
            }
        };

        return (
            <div>
                <Panel isOpen={isOpen} type={PanelType.medium} onDismiss={this._onClose.bind(this)}>
                    <div className={`bootstrap `}>
                        <div className="mb-2 modal-header">
                            <h4 className="modal-title" id="modal-title">Add Attachment</h4>
                        </div>
                        <div className="modal-body">

                            <div className="form-group">
                                <label>Attachment:</label>
                                <input type="file" multiple name="files" className={`form-control ${styles.ChooseFileButton}`} onChange={this.handleFileChange.bind(this)} />
                            </div>
                            {uploading ? <MessageBar >"Adding Attachments"</MessageBar> : null}
                            {formError ? <MessageBar messageBarType={MessageBarType.error}>{formMessage}</MessageBar> : null}
                        </div>
                        <div className="modal-footer">
                            <PrimaryButton text="Upload" disabled={disabled() || uploading || formError} onClick={this.postFiles.bind(this)}></PrimaryButton>
                            <DefaultButton text="Close" onClick={this._onClose.bind(this)}></DefaultButton>
                        </div>
                    </div>
                </Panel>
            </div>
        );
    }
}