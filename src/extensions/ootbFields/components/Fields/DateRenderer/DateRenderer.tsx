import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';
import { IFieldRendererProps } from '../Common/IFieldRendererProps';
import BaseTextRenderer from '../BaseTextRenderer/BaseTextRenderer';

export interface IDateRendererProps extends IFieldRendererProps {
    /**
     * text to be rendered
     */
    text?: string;
}

/**
 * For future
 */
export interface IDateRendererState {

}

/**
 * Field Date Renderer.
 * Used for:
 *   - Date Time
 */
export default class DateRenderer extends React.Component<IDateRendererProps, IDateRendererState> {
    public constructor(props: IDateRendererProps, state: IDateRendererState) {
        super(props, state);

        this.state = {};
    }

    @override
    public render(): JSX.Element {
            return (<BaseTextRenderer cssProps={this.props.cssProps} className={css(this.props.className)} noTextRender={true}>{this.props.text}</BaseTextRenderer>);
    }
}