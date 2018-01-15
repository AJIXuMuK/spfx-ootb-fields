import { IContext, IRegionalSettings } from '../common/Interfaces';
import { GeneralHelper } from './GeneralHelper';
import { ISPFieldLookupValue, IPrincipal, ITerm } from '../common/SPEntities';
import { FieldNamesMapping } from '../common/Constants';
import * as Constants from '../common/Constants';
import { ListItemAccessor } from '@microsoft/sp-listview-extensibility';
import { SPField } from '@microsoft/sp-page-context';
import { SPHttpClient } from '@microsoft/sp-http';

declare var SP: any;
declare var window: any;

/**
 * Helper class to work with SharePoint objects and entities
 */
export class SPHelper {
    private static _isInitialized: boolean = false;
    private static _regionalSettings: IRegionalSettings;

    /**
     * Gets field's Real Name from FieldNamesMapping
     * @param columnName current field name
     */
    public static getStoredFieldName(columnName: string): string {
        if (!columnName)
            return '';

        return FieldNamesMapping[columnName] ? FieldNamesMapping[columnName].storedName : columnName;
    }

    /**
     * Gets Field's text
     * @param fieldValue field value as it appears in Field Customizer's onRenderCell event
     * @param listItem List Item accessor
     * @param context Customizer's context
     */
    public static getFieldText(fieldValue: any, listItem: ListItemAccessor, context: IContext): string {
        const field: SPField = context.field;

        if (!field) {
            return '';
        }

        const fieldName: string = this.getFieldNameById(field.id.toString());
        const fieldType: string = field.fieldType;
        const strFieldValue: string = fieldValue ? fieldValue.toString() : '';

        switch (fieldType) {
            case 'Note':
                const isRichText: boolean = SPHelper.getFieldProperty(field.id.toString(), "RichText") === 'TRUE';
                if (isRichText) {
                    return GeneralHelper.getTextFromHTML(strFieldValue);
                }
                return fieldValue;
            case 'DateTime':
                const friendlyDisplay: string = SPHelper.getRowItemValueByName(listItem, `${fieldName}.FriendlyDisplay`);
                return friendlyDisplay ? GeneralHelper.getRelativeDateTimeString(friendlyDisplay) : strFieldValue;
            case 'User':
            case 'UserMulti':
                const titles: string[] = [];
                const users: IPrincipal[] = <IPrincipal[]>fieldValue;

                if (!users) {
                    return '';
                }

                for (let i = 0, len = users.length; i < len; i++) {
                    titles.push(users[i].title);
                }
                return titles.join('\n');
            case "Lookup":
            case "LookupMulti":
                const lookupValues = fieldValue as ISPFieldLookupValue[];

                if (!lookupValues) {
                    return '';
                }

                const lookupTexts: string[] = [];
                for (let i = 0, len = lookupValues.length; i < len; i++) {
                    lookupTexts.push(lookupValues[i].lookupValue);
                }
                return lookupTexts.join('\n');
            case 'URL':
                const isImage: boolean = SPHelper.getFieldProperty(field.id.toString(), 'Format') === 'Image';
                if (isImage) {
                    return '';
                }
                return SPHelper.getRowItemValueByName(listItem, `${fieldName}.desc`);
            case 'Taxonomy':
            case 'TaxonomyFieldType':
                const terms: ITerm[] = Array.isArray(fieldValue) ? <ITerm[]>fieldValue : <ITerm[]>[fieldValue];

                if (!terms) {
                    return '';
                }

                const termTexts: string[] = [];
                for (let i = 0, len = terms.length; i < len; i++) {
                    termTexts.push(terms[i].Label);
                }
                return termTexts.join('\n');
            case 'Attachments':
                return '';
            case 'Computed':
                const storedName: string = this.getStoredFieldName(fieldName);
                if (storedName === 'URL') {
                    return this.getRowItemValueByName(listItem, 'URL.desc') || strFieldValue;
                }
                return strFieldValue;
            default:
                return strFieldValue;
        }
    }

    /**
     * Returns Field's name by its ID
     * @param fieldId Field's ID
     * @param returnStoredName true if needs to return RealFieldName
     */
    public static getFieldNameById(fieldId: string, returnStoredName: boolean = false): string {
        // TODO: g_listData may disappear
        if (window.g_listData.ListSchema.Field) {
            const fields: any[] = window.g_listData.ListSchema.Field;

            for (let i = 0, len = fields.length; i < len; i++) {
                const field: any = fields[i];

                if (field.ID === fieldId) {
                    return returnStoredName ? field.RealFieldName : field.Name;
                }
            }
        }

        return '';
    }

    /**
     * Gets property of the Field by Field's ID and Property Name
     * @param fieldId Field's ID
     * @param propertyName Property name
     */
    public static getFieldProperty(fieldId: string, propertyName: string): any {
        // TODO: g_listData may disappear
        if (window.g_listData.ListSchema.Field) {
            const fields: any[] = window.g_listData.ListSchema.Field;

            for (let i = 0, len = fields.length; i < len; i++) {
                const field: any = fields[i];

                if (field.ID === fieldId) {
                    return field[propertyName];
                }
            }
        }

        return null;
    }

    /**
     * Gets column's value for the row by row's ID.
     * This method works with g_listData to be able to get such values as FriendlyDisplay text for Date, and more.
     * @param id row ID (item ID)
     * @param itemName column name
     */
    public static getRowItemValueById(id: string, itemName: string): any {
        // TODO: g_listData may disappear
        if (window.g_listData.ListData.Row) {
            const rows: any[] = window.g_listData.ListData.Row;

            for (let i = 0, len = rows.length; i < len; i++) {
                const row: any = rows[i];
                if (row.ID === id) {
                    return row[itemName];
                }
            }
        }

        return null;
    }

    /**
     * Gets column's value for the row using List Item Accessor.
     * This method works with private property _values of List Item Accessor to get such values as FriendlyDisplay text for Date, and more.
     * @param listItem List Item Accessor
     * @param itemName column name
     */
    public static getRowItemValueByName(listItem: ListItemAccessor, itemName: string): any {
        return (<any>listItem)._values ? ((<any>listItem)._values as Map<string, any>).get(itemName) :
            this.getRowItemValueById(listItem.getValueByName('ID'), itemName);
    }

    /**
     * Gets SchemaXml for the field by List Title and Field Internal Name
     * @param fieldName Field's Internal Name
     * @param listTitle List Title
     * @param context Customizer's context
     */
    public static getFieldSchemaXmlByInternalNameOrTitle(fieldName: string, listTitle: string, context: IContext): Promise<string> {
        return new Promise<string>((resolve) => {
            context.spHttpClient.get(`${GeneralHelper.trimSlash(context.pageContext.web.absoluteUrl)}/_api/web/lists/getByTitle('${listTitle}')/fields/getByInternalNameOrTitle('${fieldName}')?$select=SchemaXml`,
                SPHttpClient.configurations.v1).then((response) => {
                    return response.json();
                }).then(result => {
                    resolve(result ? result.SchemaXml : '');
                });
        });
    }
}