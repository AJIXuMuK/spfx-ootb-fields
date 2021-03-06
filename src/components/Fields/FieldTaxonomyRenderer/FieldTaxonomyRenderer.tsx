import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css } from 'office-ui-fabric-react';

import { ITerm } from '../../../common/SPEntities';
import { IFieldRendererProps } from '../FieldCommon/IFieldRendererProps';

import styles from './FieldTaxonomyRenderer.module.scss';

export interface IFieldTaxonomyRendererProps extends IFieldRendererProps {
    /**
     * terms to display
     */
    terms: ITerm[];
}

/**
 * For future
 */
export interface IFieldTaxonomyRendererState {

}

/**
 * Field Taxonomy Renderer.
 * Used for:
 *   - Taxonomy
 */
export default class FieldTaxonomyRenderer extends React.Component<IFieldTaxonomyRendererProps, IFieldTaxonomyRendererState> {
    public constructor(props: IFieldTaxonomyRendererProps, state: IFieldTaxonomyRendererState) {
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