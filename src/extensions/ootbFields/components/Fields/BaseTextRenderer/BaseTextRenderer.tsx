import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';

import styles from './BaseTextRenderer.module.scss';
import { IFieldRendererProps } from '../Common/IFieldRendererProps';

export interface IBaseTextRendererProps extends IFieldRendererProps {
    /**
     * text to be displayed
     */
    text?: string;
    /**
     * true if no need to render span element with text content
     */
    noTextRender?: boolean;
}

/**
 * For future
 */
export interface IBaseTextRendererState {

}

/**
 * Base renderer. Used to render text.
 */
export default class BaseTextRenderer extends React.Component<IBaseTextRendererProps, IBaseTextRendererState> {
    public constructor (props: IBaseTextRendererProps, state: IBaseTextRendererState) {
        super(props, state);

        this.state = {};
    }

    @override
    public render(): JSX.Element {
        const text: string = this.props.text || ' ';
        return (<div className={css(this.props.className, styles.baseText)} style={this.props.cssProps}>
        { this.props.noTextRender ? null : <span>{text}</span> }
        {this.props.children}
        </div>);
    }
}