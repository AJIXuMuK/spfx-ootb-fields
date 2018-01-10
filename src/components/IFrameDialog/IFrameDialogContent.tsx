import * as React from "react";
import { DialogContent } from "office-ui-fabric-react";

import styles from './IFrameDialogContent.module.scss';

export interface IIFrameDialogContentProps {
    url: string;
    close: () => void;
    iframeOnLoad?: (iframe: any) => {};
}

/**
 * IFrame Dialog content
 */
export default class IFrameDialogContent extends React.Component<IIFrameDialogContentProps, {}> {
    private _iframe: any;

    constructor (props: IIFrameDialogContentProps) {
        super(props);
    }

    public render(): JSX.Element {
        return <DialogContent
            showCloseButton={true}
            onDismiss={this.props.close}>
                <div className={styles.iFrameDialog}>
                    <iframe ref={(iframe) => { this._iframe = iframe; }} frameBorder={0} src={this.props.url} onLoad={this._iframeOnLoad.bind(this)} style={{ width: '100%', height: '315px', visibility: 'hidden' }} />
                </div>
            </DialogContent>;
    }

    private _iframeOnLoad(): void {
        this._iframe.contentWindow.frameElement.cancelPopUp = this.props.close;

        if (this.props.iframeOnLoad) {
            this.props.iframeOnLoad(this._iframe);
        }

        this._iframe.style.visibility = 'visible';
    }
}