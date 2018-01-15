import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'OotbFieldsStrings';
import OotbFields, { IOotbFieldsProps } from './components/Customizer/OotbFields';
import { SPHelper } from '../../utilities/SPHelper';
import { Promise } from 'es6-promise';

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IOotbFieldsFieldCustomizerProperties {
}

const LOG_SOURCE: string = 'OotbFieldsFieldCustomizer';

export default class OotbFieldsFieldCustomizer
  extends BaseFieldCustomizer<IOotbFieldsFieldCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated OotbFieldsFieldCustomizer with properties:');
    return Promise.resolve();
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const fieldName: string = SPHelper.getStoredFieldName(this.context.field.internalName);
    const text: string = SPHelper.getFieldText(event.fieldValue, event.listItem, this.context);

    const ootbFields: React.ReactElement<{}> =
      React.createElement(OotbFields, {
        text: text,
        value: event.fieldValue,
        listItem: event.listItem,
        fieldName: fieldName,
        context: this.context,
        cssProps: { backgroundColor: '#f00' },
        className: 'fake-class'
      });

    ReactDOM.render(ootbFields, event.domElement);
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
