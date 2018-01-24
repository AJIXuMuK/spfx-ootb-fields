import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';

import { Link } from 'office-ui-fabric-react';

import FieldBaseTextRenderer from '../FieldBaseTextRenderer/FieldBaseTextRenderer';
import { IFieldRendererProps } from '../FieldCommon/IFieldRendererProps';

export interface IFieldTitleRendererProps extends IFieldRendererProps {
    /**
     * text to be displayed
     */
    text?: string;
    /**
     * true if the Title should be rendered as link
     */
    isLink?: boolean;
    /**
     * web url
     */
    baseUrl?: string;
    /**
     * list id
     */
    listId?: string;
    /**
     * item id
     */
    id?: number;
    /**
     * custom title click event handler. If not set Display form for the item will be displaed
     */
    onClick?: (args: IFieldTitleClickEventArgs) => {};
}

/**
 * For future
 */
export interface IFieldTitleRendererState {

}

/**
 * Title click event arguments
 */
export interface IFieldTitleClickEventArgs {
    listId?: string;
    id?: string;
}

/**
 * Field Title Renderer.
 * Used for:
 *   - Title
 */
export default class FieldTitleRenderer extends React.Component<IFieldTitleRendererProps, IFieldTitleRendererState> {
    public constructor(props: IFieldTitleRendererProps, state: IFieldTitleRendererState) {
        super(props, state);

        this.state = {};
    }

    @override
    public render(): JSX.Element {
        const isLink: boolean = this.props.isLink;
        
        if (isLink) {
            return (<Link onClick={this._onClick.bind(this)} className={css(this.props.className)} style={this.props.cssProps}>{this.props.text}</Link>);
        }
        else {
            return (<FieldBaseTextRenderer className={this.props.className} cssProps={this.props.cssProps} text={this.props.text} />);
        }
    }

    private _onClick(): void {
        if (this.props.onClick) {
            const args: IFieldTitleClickEventArgs = this.props as IFieldTitleClickEventArgs;
            this.props.onClick(args);
            return;
        }
        const url: string = `${this.props.baseUrl}/_layouts/15/listform.aspx?PageType=4&ListId=${this.props.listId}&ID=${this.props.id}`;
        location.href = url;
    }
}