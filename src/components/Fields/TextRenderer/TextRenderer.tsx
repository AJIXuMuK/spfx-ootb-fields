import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';

import styles from './TextRenderer.module.scss';

import BaseTextRenderer from '../BaseTextRenderer/BaseTextRenderer';
import { IFieldRendererProps } from '../Common/IFieldRendererProps';

export interface ITextRendererProps extends IFieldRendererProps {
    /**
     * text to be displayed
     */
    text?: string;
    /**
     * true if props.text can be inserted as innerHTML of the component
     */
    isSafeForInnerHTML?: boolean;
    /**
     * true if the text should be truncated
     */
    isTruncated?: boolean;
}

/**
 * For future
 */
export interface ITextRendererState {

}

/**
 * Field Text Renderer.
 * Used for:
 *   - Single line of text
 *   - Multiline text
 *   - Choice
 *   - Checkbox
 *   - Number
 *   - Currency
 */
export default class TextRenderer extends React.Component<ITextRendererProps, ITextRendererState> {
    public constructor(props: ITextRendererProps, state: ITextRendererState) {
        super(props, state);

        this.state = {};
    }

    @override
    public render(): JSX.Element {
        const isSafeForInnerHTML: boolean = this.props.isSafeForInnerHTML;
        const isTruncatedClassNameObj: any = {};
        isTruncatedClassNameObj[styles.isTruncated] = this.props.isTruncated;
        let text: string = this.props.text;
        if (isSafeForInnerHTML && this.props.isTruncated) {
            text += `<div class=${styles.truncate} style="background: linear-gradient(to bottom, transparent, ${this.props.cssProps.background || this.props.cssProps.backgroundColor || '#ffffff'} 100%)"></div>`;
        }


        if (isSafeForInnerHTML) {
            return (<div className={css(this.props.className, styles.fieldRendererText, isTruncatedClassNameObj)} style={this.props.cssProps} dangerouslySetInnerHTML={{__html: text}}></div>);
        }
        else {
            return (<BaseTextRenderer className={css(this.props.className, styles.fieldRendererText)} cssProps={this.props.cssProps} text={this.props.text} />);
        }
    }
}