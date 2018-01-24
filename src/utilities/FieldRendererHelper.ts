import * as React from 'react';
import { ISPFieldLookupValue, ITerm, IPrincipal } from '../common/SPEntities';
import FieldTextRenderer from '../components/Fields/FieldTextRenderer/FieldTextRenderer';
import FieldDateRenderer from '../components/Fields/FieldDateRenderer/FieldDateRenderer';
import { ListItemAccessor } from '@microsoft/sp-listview-extensibility';
import { SPHelper } from './SPHelper';
import FieldTitleRenderer from '../components/Fields/FieldTitleRenderer/FieldTitleRenderer';
import { SPField } from '@microsoft/sp-page-context';
import { IContext } from '../common/Interfaces';
import { GeneralHelper } from './GeneralHelper';
import FieldLookupRenderer, { IFieldLookupClickEventArgs } from '../components/Fields/FieldLookupRenderer/FieldLookupRenderer';
import IFrameDialog from '../components/IFrameDialog/IFrameDialog';
import FieldUrlRenderer from '../components/Fields/FieldUrlRenderer/FieldUrlRenderer';
import FieldTaxonomyRenderer from '../components/Fields/FieldTaxonomyRenderer/FieldTaxonomyRenderer';
import { IFieldRendererProps } from '../components/Fields/FieldCommon/IFieldRendererProps';
import FieldUserRenderer from '../components/Fields/FieldUserRenderer/FieldUserRenderer';
import FieldFileTypeRenderer from '../components/Fields/FieldFileTypeRenderer/FieldFileTypeRenderer';
import FieldAttachmentsRenderer from '../components/Fields/FieldAttachmentsRenderer/FieldAttachmentsRenderer';
import FieldNameRenderer from '../components/Fields/FieldNameRenderer/FieldNameRenderer';

/**
 * Field Renderer Helper.
 * Helps to render fields similarly to OOTB SharePoint rendering
 */
export class FieldRendererHelper {
    /**
     * Returns JSX.Element with OOTB rendering and applied additional props
     * @param fieldValue Value of the field
     * @param props IFieldRendererProps (CSS classes and CSS styles)
     * @param listItem Current list item
     * @param context Customizer context
     */
    public static getFieldRenderer(fieldValue: any, props: IFieldRendererProps, listItem: ListItemAccessor, context: IContext): Promise<JSX.Element> {
        return new Promise<JSX.Element>(resolve => {
            const field: SPField = context.field;
            const listId: string = context.pageContext.list.id.toString();
            const fieldType: string = field.fieldType;
            const fieldName: string = field.internalName; //SPHelper.getFieldNameById(field.id.toString());
            let result: JSX.Element = null;
            const fieldValueAsEncodedText: string = fieldValue ? GeneralHelper.encodeText(fieldValue.toString()) : '';

            switch (fieldType) {
                case 'Text':
                case 'Choice':
                case 'Boolean':
                case 'MultiChoice':
                case 'Computed':
                    const fieldStoredName: string = SPHelper.getStoredFieldName(fieldName);
                    if (fieldStoredName === 'Title') {
                        resolve(React.createElement(FieldTitleRenderer, {
                            text: fieldValueAsEncodedText,
                            isLink: fieldName === 'LinkTitle' || fieldName === 'LinkTitleNoMenu',
                            listId: listId,
                            id: listItem.getValueByName('ID'),
                            baseUrl: context.pageContext.web.absoluteUrl,
                            ...props
                        }));
                    }
                    else if (fieldStoredName === 'DocIcon') {
                        const path: string = listItem.getValueByName('FileLeafRef');
                        resolve(React.createElement(FieldFileTypeRenderer, {
                            path: path,
                            isFolder: SPHelper.getRowItemValueByName(listItem, 'FSObjType') === '1',
                            ...props
                        }));
                    }
                    else if (fieldStoredName === 'FileLeafRef') {
                        resolve(React.createElement(FieldNameRenderer, {
                            text: fieldValueAsEncodedText,
                            isLink: true,
                            filePath: SPHelper.getRowItemValueByName(listItem, 'FileRef'),
                            isNew: SPHelper.getRowItemValueByName(listItem, 'Created_x0020_Date.ifnew') === '1',
                            hasPreview: true,
                            ...props
                        }));
                    }
                    else if (fieldStoredName === 'URL') {
                        resolve(React.createElement(FieldUrlRenderer, {
                            isImageUrl: false,
                            url: fieldValue.toString(),
                            text: SPHelper.getRowItemValueByName(listItem, `URL.desc`) || fieldValueAsEncodedText,
                            ...props
                        }));
                    }
                    else {
                        resolve(React.createElement(FieldTextRenderer, {
                            text: fieldValueAsEncodedText,
                            isSafeForInnerHTML: false,
                            isTruncated: false,
                            ...props
                        }));
                    }
                    break;
                case 'Integer':
                case 'Counter':
                case 'Number':
                case 'Currency':
                    resolve(React.createElement(FieldTextRenderer, {
                        text: fieldValueAsEncodedText,
                        isSafeForInnerHTML: true,
                        isTruncated: false,
                        ...props
                    }));
                    break;
                case 'Note':
                    SPHelper.getFieldProperty(field.id.toString(), "RichText", context, false).then(richText => {
                        const isRichText: boolean = richText === 'TRUE';
                        let html: string = '';

                        if (isRichText) {
                            html = fieldValue.toString();
                        }
                        else {
                            html = fieldValueAsEncodedText.replace(/\n/g, "<br>");
                        }
                        // text is truncated if its length is more that 255 symbols or it has more than 4 lines
                        let isTruncated: boolean = html.length > 255 || html.split(/\r|\r\n|\n|<br>/).length > 4;
                        resolve(React.createElement(FieldTextRenderer, {
                            text: html,
                            isSafeForInnerHTML: true,
                            isTruncated: isTruncated,
                            ...props
                        }));
                    });
                    break;
                case 'DateTime':
                    const friendlyDisplay: string = SPHelper.getRowItemValueByName(listItem, `${fieldName}.FriendlyDisplay`);
                    resolve(React.createElement(FieldDateRenderer, {
                        text: friendlyDisplay ? GeneralHelper.getRelativeDateTimeString(friendlyDisplay) : fieldValueAsEncodedText,
                        ...props
                    }));
                    break;
                case "Lookup":
                case "LookupMulti":
                    SPHelper.getLookupFieldListDispFormUrl(field.id.toString(), context).then(dispFormUrlValue => {
                        const lookupValues = fieldValue as ISPFieldLookupValue[];
                        const dispFormUrl: string = dispFormUrlValue.toString();
                        resolve(React.createElement(FieldLookupRenderer, {
                            lookups: lookupValues,
                            dispFormUrl: dispFormUrl,
                            ...props
                        }));
                    });

                    break;
                case 'URL':
                    SPHelper.getFieldProperty(field.id.toString(), 'Format', context, true).then(format => {
                        const isImage: boolean = format === 'Image';
                        const text: string = SPHelper.getRowItemValueByName(listItem, `${fieldName}.desc`);
                        resolve(React.createElement(FieldUrlRenderer, {
                            isImageUrl: isImage,
                            url: fieldValue.toString(),
                            text: text,
                            ...props
                        }));
                    });
                    break;
                case 'Taxonomy':
                case 'TaxonomyFieldType':
                case 'TaxonomyFieldTypeMulti':
                    const terms: ITerm[] = Array.isArray(fieldValue) ? <ITerm[]>fieldValue : <ITerm[]>[fieldValue];
                    resolve(React.createElement(FieldTaxonomyRenderer, {
                        terms: terms,
                        ...props
                    }));
                    break;
                case 'User':
                case 'UserMulti':
                    resolve(React.createElement(FieldUserRenderer, {
                        users: <IPrincipal[]>fieldValue,
                        context: context,
                        ...props
                    }));
                    break;
                case 'Attachments':
                    resolve(React.createElement(FieldAttachmentsRenderer, {
                        count: parseInt(fieldValue),
                        ...props
                    }));
                    break;
                default:
                    resolve(React.createElement(FieldTextRenderer, {
                        text: fieldValueAsEncodedText,
                        isSafeForInnerHTML: false,
                        isTruncated: false,
                        ...props
                    }));
                    break;
            }
        });
    }
}