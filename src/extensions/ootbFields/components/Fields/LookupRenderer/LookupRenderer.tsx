import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';
import { Link } from 'office-ui-fabric-react';

import { ISPFieldLookupValue } from "../../../common/SPEntities";
import { IFieldRendererProps } from '../Common/IFieldRendererProps';

import styles from './LookupRenderer.module.scss';
import IFrameDialog from '../../IFrameDialog/IFrameDialog';

export interface ILookupRendererProps extends IFieldRendererProps {
    /**
     * lookup values
     */
    lookups: ISPFieldLookupValue[];
    /**
     * url of Display form for the list that is referenced in the lookup
     */
    dispFormUrl?: string;
    /**
     * custom event handler of lookup item click. If not set the dialog with Display Form will be shown
     */
    onClick?: (args: ILookupClickEventArgs) => {};
}

/**
 * For future
 */
export interface ILookupRendererState {

}

/**
 * Lookup click event arguments
 */
export interface ILookupClickEventArgs {
    lookup?: ISPFieldLookupValue;
}

/**
 * Field Lookup Renderer.
 * Used for:
 *   - Lookup, LookupMulti
 */
export default class LookupRenderer extends React.Component<ILookupRendererProps, ILookupRendererState> {
    public constructor(props: ILookupRendererProps, state: ILookupRendererState) {
        super(props, state);

        this.state = {};
    }

    @override
    public render(): JSX.Element {
        const lookupLinks: JSX.Element[] = this.props.lookups.map((lookup) => {
            return <Link onClick={this._onClick.bind(this, lookup.lookupId)} className={styles.lookup} style={this.props.cssProps}>{lookup.lookupValue}</Link>;
        });
        return (<div style={this.props.cssProps} className={css(this.props.className)}>{lookupLinks}</div>);
    }

    private _onClick(lookup: ISPFieldLookupValue): void {
        if (this.props.onClick) {
            const args: ILookupClickEventArgs = {
                lookup: lookup
            };
            this.props.onClick(args);
            return;
        }
        
        //
        // showing Display Form in the dialog
        //
        const iFrameDlg: IFrameDialog = new IFrameDialog();
        iFrameDlg.url = `${this.props.dispFormUrl}&ID=${lookup.lookupId}&RootFolder=*&IsDlg=1`;
        iFrameDlg.iframeOnLoad = this._onIframeLoaded.bind(this);
        iFrameDlg.show();
    }

    private _onIframeLoaded(iframe: any): void {
        //
        // some additional configuration to beutify content of the iframe
        //
        const iframeWindow: Window = iframe.contentWindow;
        const iframeDocument: Document = iframeWindow.document;

        const s4Workspace: HTMLDivElement = iframeDocument.getElementById('s4-workspace') as HTMLDivElement;
        s4Workspace.style.height = iframe.style.height;
        s4Workspace.scrollIntoView();
    }
}