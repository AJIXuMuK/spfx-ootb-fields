import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';

import { Link, Icon } from 'office-ui-fabric-react';

import BaseTextRenderer from '../BaseTextRenderer/BaseTextRenderer';
import { IFieldRendererProps } from '../Common/IFieldRendererProps';

import styles from './NameRenderer.module.scss';
import { GeneralHelper } from "../../../utilities/GeneralHelper";

export interface INameRendererProps extends IFieldRendererProps {
    /**
     * text to display
     */
    text?: string;
    /**
     * if the Name should be rendered as link
     */
    isLink: boolean;
    /**
     * path to the document
     */
    filePath?: string;
    /**
     * true if the document is new
     */
    isNew?: boolean;
    /**
     * true if the document type has preview (true by default)
     */
    hasPreview?: boolean;
    /**
     * custom handler for link click. If not set link click will lead to rendering document preview
     */
    onClick?: (args: INameClickEventArgs) => {};
}

/**
 * For future
 */
export interface INameRendererState {

}

/**
 * Name click event arguments
 */
export interface INameClickEventArgs {
    filePath?: string;
}

/**
 * Field Title Renderer.
 * Used for:
 *   - Title
 */
export default class NameRenderer extends React.Component<INameRendererProps, INameRendererState> {
    public constructor(props: INameRendererProps, state: INameRendererState) {
        super(props, state);

        this.state = {};
    }

    @override
    public render(): JSX.Element {
        const isLink: boolean = this.props.isLink;
        //
        // for now only signal for New documents is implemented
        //
        let signal: JSX.Element = this.props.isNew ? <span className={css(styles.signal, styles.newItem)}><Icon iconName={'Glimmer'} className={css(styles.newIcon)}/></span> : null;
        let value: JSX.Element;

        if (isLink) {
            if (this.props.onClick) {
                value = <Link onClick={this._onClick.bind(this)} style={this.props.cssProps} className={styles.value}>{this.props.text}</Link>;
            }
            else {
                let url: string;
                const filePath = this.props.filePath;
                if (this.props.hasPreview !== false) {
                    url = `#id=${filePath}&parent=${filePath.substring(0, filePath.lastIndexOf('/'))}`;
                }
                else {
                    url = filePath;
                }

                value = <Link href={url} style={this.props.cssProps} className={styles.value}>{this.props.text}</Link>;
            }
        }
        else {
            value = <BaseTextRenderer cssProps={this.props.cssProps} text={this.props.text} />;
        }

        return <span className={css(styles.signalField, this.props.className)} style={this.props.cssProps}>
            {signal}
            <span className={styles.signalFieldValue}>
                <span data-selection-invoke={'true'}>
                    {value}
                </span>
            </span>
        </span>;
    }

    private _onClick(): void {
        if (this.props.onClick) {
            const args: INameClickEventArgs = this.props as INameClickEventArgs;
            this.props.onClick(args);
            return;
        }
    }
}