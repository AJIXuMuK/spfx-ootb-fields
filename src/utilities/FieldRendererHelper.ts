import * as React from 'react';
import { ISPFieldLookupValue, ITerm, IPrincipal } from '../common/SPEntities';
import TextRenderer from '../components/Fields/TextRenderer/TextRenderer';
import DateRenderer from '../components/Fields/DateRenderer/DateRenderer';
import { ListItemAccessor } from '@microsoft/sp-listview-extensibility';
import { SPHelper } from './SPHelper';
import TitleRenderer from '../components/Fields/TitleRenderer/TitleRenderer';
import { SPField } from '@microsoft/sp-page-context';
import { IContext } from '../common/Interfaces';
import { GeneralHelper } from './GeneralHelper';
import LookupRenderer, { ILookupClickEventArgs } from '../components/Fields/LookupRenderer/LookupRenderer';
import IFrameDialog from '../components/IFrameDialog/IFrameDialog';
import UrlRenderer from '../components/Fields/UrlRenderer/UrlRenderer';
import TaxonomyRenderer from '../components/Fields/TaxonomyRenderer/TaxonomyRenderer';
import { IFieldRendererProps } from '../components/Fields/Common/IFieldRendererProps';
import UserRenderer from '../components/Fields/UserRenderer/UserRenderer';
import FileTypeRenderer from '../components/Fields/FileTypeRenderer/FileTypeRenderer';
import AttachmentsRenderer from '../components/Fields/AttachmentsRenderer/AttachmentsRenderer';
import NameRenderer from '../components/Fields/NameRenderer/NameRenderer';

declare var SP: any;

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
    public static getFieldRenderer(fieldValue: any, props: IFieldRendererProps, listItem: ListItemAccessor, context: IContext): JSX.Element {
        const field: SPField = context.field;
        const listId: string = context.pageContext.list.id.toString();
        const fieldType: string = field.fieldType;
        const fieldName: string = SPHelper.getFieldNameById(field.id.toString());
        let result: JSX.Element = null;
        const fieldValueAsEncodedText: string = fieldValue ? GeneralHelper.encodeText(fieldValue.toString()) : '';

        switch (fieldType) {
            case 'Text':
            case 'Choice':
            case 'Boolean':
            case 'MultiChoice':
                result = React.createElement(TextRenderer, {
                    text: fieldValueAsEncodedText,
                    isSafeForInnerHTML: false,
                    isTruncated: false,
                    ...props
                });
                break;
            case 'Computed':
                const fieldStoredName: string = SPHelper.getStoredFieldName(fieldName);
                if (fieldStoredName === 'Title') {
                    result = React.createElement(TitleRenderer, {
                        text: fieldValueAsEncodedText,
                        isLink: fieldName === 'LinkTitle' || fieldName === 'LinkTitleNoMenu',
                        listId: listId,
                        id: listItem.getValueByName('ID'),
                        baseUrl: context.pageContext.web.absoluteUrl,
                        ...props
                    });
                }
                else if (fieldStoredName === 'DocIcon') {
                    const path: string = listItem.getValueByName('FileLeafRef');
                    result = React.createElement(FileTypeRenderer, {
                        path: path,
                        isFolder: SPHelper.getRowItemValueByName(listItem, 'FSObjType') === '1',
                        ...props
                    });
                }
                else if (fieldStoredName === 'FileLeafRef') {
                    result = React.createElement(NameRenderer, {
                        text: fieldValueAsEncodedText,
                        isLink: true,
                        filePath: SPHelper.getRowItemValueByName(listItem, 'FileRef'),
                        isNew: SPHelper.getRowItemValueByName(listItem, 'Created_x0020_Date.ifnew') === '1',
                        hasPreview: true,
                        ...props
                    });
                }
                else if (fieldStoredName === 'URL') {
                    result = React.createElement(UrlRenderer, {
                        isImageUrl: false,
                        url: fieldValue.toString(),
                        text: SPHelper.getRowItemValueByName(listItem, `URL.desc`) || fieldValueAsEncodedText,
                        ...props
                    });
                }
                else {
                    result = React.createElement(TextRenderer, {
                        text: fieldValueAsEncodedText,
                        isSafeForInnerHTML: false,
                        isTruncated: false,
                        ...props
                    });
                }
                break;
            case 'Integer':
            case 'Counter':
            case 'Number':
            case 'Currency':
                result = React.createElement(TextRenderer, {
                    text: fieldValueAsEncodedText,
                    isSafeForInnerHTML: true,
                    isTruncated: false,
                    ...props
                });
                break;
            case 'Note':
                const isRichText: boolean = SPHelper.getFieldProperty(field.id.toString(), "RichText") === 'TRUE';
                let html: string = '';

                if (isRichText) {
                    html = fieldValue.toString();
                }
                else {
                    html = fieldValueAsEncodedText.replace(/\n/g, "<br>");
                }
                // text is truncated if its length is more that 255 symbols or it has more than 4 lines
                let isTruncated: boolean = html.length > 255 || html.split(/\r|\r\n|\n|<br>/).length > 4;
                result = React.createElement(TextRenderer, {
                    text: html,
                    isSafeForInnerHTML: true,
                    isTruncated: isTruncated,
                    ...props
                });
                break;
            case 'DateTime':
                const friendlyDisplay: string = SPHelper.getRowItemValueByName(listItem, `${fieldName}.FriendlyDisplay`);
                result = React.createElement(DateRenderer, {
                    text: friendlyDisplay ? GeneralHelper.getRelativeDateTimeString(friendlyDisplay) : fieldValueAsEncodedText,
                    ...props
                });
                break;
            case "Lookup":
            case "LookupMulti":
                const lookupValues = fieldValue as ISPFieldLookupValue[];
                const dispFormUrl: string = SPHelper.getFieldProperty(field.id.toString(), 'DispFormUrl').toString();
                result = React.createElement(LookupRenderer, {
                    lookups: lookupValues,
                    dispFormUrl: dispFormUrl,
                    ...props
                });
                break;
            case 'URL':
                const isImage: boolean = SPHelper.getFieldProperty(field.id.toString(), 'Format') === 'Image';
                const text: string = SPHelper.getRowItemValueByName(listItem, `${fieldName}.desc`);
                result = React.createElement(UrlRenderer, {
                    isImageUrl: isImage,
                    url: fieldValue.toString(),
                    text: text,
                    ...props
                });
                break;
            case 'Taxonomy':
            case 'TaxonomyFieldType':
                const terms: ITerm[] = Array.isArray(fieldValue) ? <ITerm[]>fieldValue : <ITerm[]>[fieldValue];
                result = React.createElement(TaxonomyRenderer, {
                    terms: terms,
                    ...props
                });
                break;
            case 'User':
            case 'UserMulti':
                result = React.createElement(UserRenderer, {
                    users: <IPrincipal[]>fieldValue,
                    context: context,
                    ...props
                });
                break;
            case 'Attachments':
                result = React.createElement(AttachmentsRenderer, {
                    count: parseInt(fieldValue),
                    ...props
                });
                break;
            default:
                result = React.createElement(TextRenderer, {
                    text: fieldValueAsEncodedText,
                    isSafeForInnerHTML: false,
                    isTruncated: false,
                    ...props
                });
                break;
        }

        return result;
    }
}