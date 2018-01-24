import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';
import { IFieldRendererProps } from '../FieldCommon/IFieldRendererProps';
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";

import styles from './FieldAttachmentsRenderer.module.scss';

/**
 * Attachments renderer props
 */
export interface IFieldAttachmentsRendererProps extends IFieldRendererProps {
    /**
     * amount of attachments
     */
    count?: number;
}

/**
 * For future
 */
export interface IFieldAttahcmentsRendererState {

}

/**
 * Attachments Renderer.
 * Used for:
 *   - Attachments
 */
export default class FieldAttachmentsRenderer extends React.Component<IFieldAttachmentsRendererProps, IFieldAttahcmentsRendererState> {
    public constructor(props: IFieldAttachmentsRendererProps, state: IFieldAttahcmentsRendererState) {
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