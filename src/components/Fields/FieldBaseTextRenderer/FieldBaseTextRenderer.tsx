import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';

import styles from './FieldBaseTextRenderer.module.scss';
import { IFieldRendererProps } from '../FieldCommon/IFieldRendererProps';

export interface IFieldBaseTextRendererProps extends IFieldRendererProps {
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
export interface IFieldBaseTextRendererState {

}

/**
 * Base renderer. Used to render text.
 */
export default class FieldBaseTextRenderer extends React.Component<IFieldBaseTextRendererProps, IFieldBaseTextRendererState> {
    public constructor (props: IFieldBaseTextRendererProps, state: IFieldBaseTextRendererState) {
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