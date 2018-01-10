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
     * Returns cached regional settings.
     * This method must be used after getWebRegionalSettingsAsync is called
     */
    public static getWebRegionalSettings(): IRegionalSettings {
        if (this._regionalSettings) {
            return this._regionalSettings;
        }

        return {
            decimalSeparator: '.',
            thousandSeparator: ',',
            hoursOffset: 0,
            webDateFormat: 'MM/DD/YYYY',
            workDays: 62,
            firstDayOfWeek: 0,
            firstWeekOfYear: 0
        };
    }

    /**
     * Gets Web Regional Settings to be used in the application
     * @param context current context
     */
    public static getWebRegionalSettingsAsync(context: IContext): Promise<IRegionalSettings> {

        return new Promise<IRegionalSettings>((resolve) => {
            if (!this._isInitialized) {
                this.initialize();
            }

            if (this._regionalSettings) {
                resolve(this._regionalSettings);
                return;
            }

            GeneralHelper.loadJSOMScriptsInSequence(context)
                .then(() => {
                    const ctx = SP.ClientContext.get_current();
                    const web = ctx.get_web();
                    let regSettings = web.get_regionalSettings();
                    let dateString = SP.Utilities.Utility.formatDateTime(ctx, web, new Date(1999, 3, 6), SP.Utilities.DateTimeFormat.dateOnly);
                    const serverTZ = regSettings.get_timeZone();
                    const currDate = new Date();

                    currDate.setHours(currDate.getUTCHours());
                    var serverCurrDateResult = serverTZ.utcToLocalTime(currDate);
                    ctx.load(regSettings);

                    ctx.executeQueryAsync(() => {
                        const currTimeZoneHours: number = currDate.getHours();
                        const currTimeZoneDate: number = currDate.getDate();
                        const serverCurrDate: Date = serverCurrDateResult.get_value();
                        const serverTimeZoneDate: number = serverCurrDate.getDate();
                        const serverTimeZoneHours: number = serverCurrDate.getHours();
                        let hoursOffset: number = 0;

                        if (serverTimeZoneDate == currTimeZoneDate)
                            hoursOffset = serverTimeZoneHours - currTimeZoneHours;
                        else if (serverTimeZoneDate < currTimeZoneDate)
                            hoursOffset = serverTimeZoneHours - 24 - currTimeZoneHours;
                        else
                            hoursOffset = serverTimeZoneHours + 24 - currTimeZoneHours;


                        SPHelper._regionalSettings = {
                            decimalSeparator: regSettings.get_decimalSeparator(),
                            thousandSeparator: regSettings.get_thousandSeparator(),
                            hoursOffset: hoursOffset,
                            webDateFormat: SPHelper._getWebDateFormat(dateString.m_value),
                            workDays: regSettings.get_workDays(),
                            firstDayOfWeek: regSettings.get_firstDayOfWeek(),
                            firstWeekOfYear: regSettings.get_firstWeekOfYear()
                        };

                        window.sessionStorage.setItem(Constants.RegionalSettingsKey, JSON.stringify(SPHelper._regionalSettings));

                        resolve(SPHelper._regionalSettings);

                    }, () => {
                        SPHelper._regionalSettings = {
                            decimalSeparator: '.',
                            thousandSeparator: ',',
                            hoursOffset: 0,
                            webDateFormat: 'MM/DD/YYYY',
                            workDays: 62,
                            firstDayOfWeek: 0,
                            firstWeekOfYear: 0
                        };

                        window.sessionStorage.setItem(Constants.RegionalSettingsKey, JSON.stringify(SPHelper._regionalSettings));

                        resolve(SPHelper._regionalSettings);
                    });
                });
        });
    }

    /**
     * Initializes the class
     */
    public static initialize(): void {
        if (this._isInitialized) {
            return;
        }

        const sessionStorage: any = window.sessionStorage;

        const loadedRegionalSettingsString: string = sessionStorage.getItem(Constants.RegionalSettingsKey);
        if (loadedRegionalSettingsString) {
            this._regionalSettings = <IRegionalSettings>JSON.parse(loadedRegionalSettingsString);
        }

        this._isInitialized = true;
    }

    /**
     * Clears Session storage from added values
     */
    public static dispose(): void {
        const sessionStorage: any = window.sessionStorage;
        sessionStorage.removeItem(Constants.RegionalSettingsKey);
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

                for (let i = 0, len = users.length; i < len; i++) {
                    titles.push(users[i].title);
                }
                return titles.join('\n');
            case "Lookup":
            case "LookupMulti":
                const lookupValues = fieldValue as ISPFieldLookupValue[];
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
     * Gets Field's value
     * @param fieldValue field value as it appears in Field Customizer's onRenderCell event
     * @param listItem List Item accessor
     * @param context Customizer's context
     */
    public static getFieldValue(fieldValue: any, listItem: ListItemAccessor, context: IContext): any {
        const field: SPField = context.field;

        if (!field) {
            return '';
        }

        const fieldName: string = this.getFieldNameById(field.id.toString());
        const fieldType: string = field.fieldType;
        const strFieldValue: string = fieldValue ? fieldValue.toString() : '';

        switch (fieldType) {
            case 'DateTime':
                return GeneralHelper.parseDateCurrentRegionalSettings(strFieldValue);
            case 'Integer':
            case 'Counter':
            case 'Number':
            case 'Currency':
                return parseFloat(SPHelper.getRowItemValueByName(listItem, `${fieldName}.`) || fieldValue);
            default:
                return fieldValue;
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

    /**
     * Parses the string representation of predefined date to get date format
     * @param dateString string representation of April, 6 1999
     */
    private static _getWebDateFormat(dateString: string): string {
        let format: string = dateString.toString();
        format = format.replace(/9/gmi, 'Y').replace('1', 'Y').replace('04', 'MM').replace('4', 'M').replace(/\d/gmi, 'D');

        return format;
    }
}