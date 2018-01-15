import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';
import { Link } from 'office-ui-fabric-react';

import { IFieldRendererProps } from '../Common/IFieldRendererProps';

import styles from './UrlRenderer.module.scss';

export interface IUrlRendererProps extends IFieldRendererProps {
    /**
     * text to be displayed
     */
    text?: string;
    /**
     * url
     */
    url?: string;
    /**
     * if the field should be rendered as image
     */
    isImageUrl?: boolean;
}

/**
 * For future
 */
export interface IUrlRendererState {

}

/**
 * Field URL Renderer.
 * Used for:
 *   - URL (Hyperlink, Image)
 */
export default class UrlRenderer extends React.Component<IUrlRendererProps, IUrlRendererState> {
    public constructor(props: IUrlRendererProps, state: IUrlRendererState) {
        super(props, state);

        this.state = {};
    }

    @override
    public render(): JSX.Element {
        const isImageUrl: boolean = this.props.isImageUrl;
        
        if (isImageUrl) {
            return (<div className={css(this.props.className, styles.image)} style={this.props.cssProps} onClick={this._onImgClick.bind(this)}><img src={this.props.url} alt={this.props.text} /></div>);
        }
        else {
            return (<Link className={css(this.props.className, styles.link)} target={'_blank'} href={this.props.url} style={this.props.cssProps}>{this.props.text}</Link>);
        }
    }

    private _onImgClick(): void {
        window.open(this.props.url, '_blank');
    }
}