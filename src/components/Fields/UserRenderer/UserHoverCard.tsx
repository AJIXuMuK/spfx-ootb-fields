import * as React from 'react';
import { HoverCard, IExpandingCardProps } from 'office-ui-fabric-react';

import styles from './UserRenderer.module.scss';

export interface IUserHoverCardProps {
    /**
     * IExpandingCardProps
     */
    expandingCardProps: IExpandingCardProps;
    /**
     * User display name
     */
    displayName: string;
    /**
     * CSS styles to apply to the renderer
     */
    cssProps?: React.CSSProperties;
}

export interface IUserHoverCardState {
    contentRendered?: HTMLDivElement;
}

/**
 * Component to render User name with related Hover Card
 */
export default class UserHoverCard extends React.Component<IUserHoverCardProps, IUserHoverCardState> {
    constructor(props: IUserHoverCardProps) {
        super(props);

        this.state = {
            contentRendered: undefined
        };
    }

    public render(): JSX.Element {
        return (
            <div className={styles.user} style={this.props.cssProps}>
                <span ref={(c: HTMLDivElement) => !this.state.contentRendered && this.setState({ contentRendered: c })} data-is-focusable={true}>{this.props.displayName}</span>
                {this.state.contentRendered && this.props.expandingCardProps.onRenderCompactCard &&
                    <HoverCard
                        expandingCardProps={this.props.expandingCardProps}
                        target={this.state.contentRendered}
                        cardDismissDelay={0}
                        trapFocus={true}
                    />}
            </div>
        );
    }
}