import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';
import { IFieldRendererProps } from '../Common/IFieldRendererProps';
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";

import styles from './AttachmentsRenderer.module.scss';

/**
 * Attachments renderer props
 */
export interface IAttachmentsRendererProps extends IFieldRendererProps {
    /**
     * amount of attachments
     */
    count?: number;
}

/**
 * For future
 */
export interface IAttahcmentsRendererState {

}

/**
 * Attachments Renderer.
 * Used for:
 *   - Attachments
 */
export default class AttachmentsRenderer extends React.Component<IAttachmentsRendererProps, IAttahcmentsRendererState> {
    public constructor(props: IAttachmentsRendererProps, state: IAttahcmentsRendererState) {
        super(props, state);

        this.state = {};
    }

    @override
    public render(): JSX.Element {
        return (
            <div className={css(this.props.className, styles.container)} style={this.props.cssProps}>
                {this.props.count && <i className='ms-Icon ms-Icon--Attach'></i>}
            </div>
        );
    }
}