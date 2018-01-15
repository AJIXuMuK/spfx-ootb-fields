import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import * as React from 'react';

import styles from './OotbFields.module.scss';
import { ListItemAccessor } from '@microsoft/sp-listview-extensibility';
import { FieldRendererHelper } from '../../../../utilities/FieldRendererHelper';
import { IProps } from '../../../../common/Interfaces';
import { IFieldRendererProps } from '../../../../components/Fields/Common/IFieldRendererProps';

export interface IOotbFieldsProps extends IProps, IFieldRendererProps {
  text: string;
  value: any;
  listItem: ListItemAccessor;
  fieldName: string;
}

const LOG_SOURCE: string = 'OotbFields';

/**
 * Field Customizer control to test fields' controls
 */
export default class OotbFields extends React.Component<IOotbFieldsProps, {}> {
  @override
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: OotbFields mounted');
  }

  @override
  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: OotbFields unmounted');
  }

  @override
  public render(): React.ReactElement<{}> {
    return (
      <div className={styles.cell}>
        {FieldRendererHelper.getFieldRenderer(this.props.value, {
          className: this.props.className,
          cssProps: this.props.cssProps
        }, this.props.listItem, this.props.context)}
      </div>
    );
  }
}
