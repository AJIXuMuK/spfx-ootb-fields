import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';

import { ITerm } from '../../../common/SPEntities';
import { IFieldRendererProps } from '../Common/IFieldRendererProps';

import styles from './TaxonomyRenderer.module.scss';

export interface ITaxonomyRendererProps extends IFieldRendererProps {
    /**
     * terms to display
     */
    terms: ITerm[];
}

/**
 * For future
 */
export interface ITaxonomyRendererState {

}

/**
 * Field Taxonomy Renderer.
 * Used for:
 *   - Taxonomy
 */
export default class TaxonomyRenderer extends React.Component<ITaxonomyRendererProps, ITaxonomyRendererState> {
    public constructor(props: ITaxonomyRendererProps, state: ITaxonomyRendererState) {
        super(props, state);

        this.state = {};
    }

    @override
    public render(): JSX.Element {
        const termEls: JSX.Element[] = this.props.terms.map((term) => {
            return <div className={styles.term} style={this.props.cssProps}><span>{term.Label}</span></div>;
        });
        return (<div style={this.props.cssProps} className={css(this.props.className)}>{termEls}</div>);
    }
}