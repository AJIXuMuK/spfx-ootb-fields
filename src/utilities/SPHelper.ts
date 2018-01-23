import { IContext } from '../common/Interfaces';
import { GeneralHelper } from './GeneralHelper';
import { ISPFieldLookupValue, IPrincipal, ITerm } from '../common/SPEntities';
import { FieldNamesMapping } from '../common/Constants';
import * as Constants from '../common/Constants';
import { ListItemAccessor } from '@microsoft/sp-listview-extensibility';
import { SPField } from '@microsoft/sp-page-context';
import { sp } from '@pnp/sp';
import { ISPField, IFields } from '../Common';

declare var window: any;

/**
 * Helper class to work with SharePoint objects and entities
 */
export class SPHelper {

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
    public static getFieldText(fieldValue: any, listItem: ListItemAccessor, context: IContext): Promise<string> {
        return new Promise<string>(resolve => {
            const field: SPField = context.field;

            if (!field) {
                resolve('');
                return;
            }

            const fieldName: string = field.internalName; //this.getFieldNameById(field.id.toString());
            const fieldType: string = field.fieldType;
            const strFieldValue: string = fieldValue ? fieldValue.toString() : '';

            switch (fieldType) {
                case 'Note':
                    SPHelper.getFieldProperty(field.id.toString(), "RichText", context).then(richText => {
                        const isRichText: boolean = richText === 'TRUE';
                        if (isRichText) {
                            resolve(GeneralHelper.getTextFromHTML(strFieldValue));
                        }
                        resolve(fieldValue);
                    });
                    break;
                case 'DateTime':
                    const friendlyDisplay: string = SPHelper.getRowItemValueByName(listItem, `${fieldName}.FriendlyDisplay`);
                    resolve(friendlyDisplay ? GeneralHelper.getRelativeDateTimeString(friendlyDisplay) : strFieldValue);
                    break;
                case 'User':
                case 'UserMulti':
                    const titles: string[] = [];
                    const users: IPrincipal[] = <IPrincipal[]>fieldValue;

                    if (!users) {
                        resolve('');
                    }

                    for (let i = 0, len = users.length; i < len; i++) {
                        titles.push(users[i].title);
                    }
                    resolve(titles.join('\n'));
                    break;
                case "Lookup":
                case "LookupMulti":
                    const lookupValues = fieldValue as ISPFieldLookupValue[];

                    if (!lookupValues) {
                        resolve('');
                    }

                    const lookupTexts: string[] = [];
                    for (let i = 0, len = lookupValues.length; i < len; i++) {
                        lookupTexts.push(lookupValues[i].lookupValue);
                    }
                    resolve(lookupTexts.join('\n'));
                    break;
                case 'URL':
                    SPHelper.getFieldProperty(field.id.toString(), 'Format', context).then(format => {
                        const isImage: boolean = format === 'Image';
                        if (isImage) {
                            resolve('');
                        }
                        resolve(SPHelper.getRowItemValueByName(listItem, `${fieldName}.desc`));
                    });
                    break;
                case 'Taxonomy':
                case 'TaxonomyFieldType':
                case 'TaxonomyFieldTypeMulti':
                    const terms: ITerm[] = Array.isArray(fieldValue) ? <ITerm[]>fieldValue : <ITerm[]>[fieldValue];

                    if (!terms) {
                        resolve('');
                    }

                    const termTexts: string[] = [];
                    for (let i = 0, len = terms.length; i < len; i++) {
                        termTexts.push(terms[i].Label);
                    }
                    resolve(termTexts.join('\n'));
                    break;
                case 'Attachments':
                    resolve('');
                    break;
                case 'Computed':
                    const storedName: string = this.getStoredFieldName(fieldName);
                    if (storedName === 'URL') {
                        resolve(this.getRowItemValueByName(listItem, 'URL.desc') || strFieldValue);
                    }
                    resolve(strFieldValue);
                    break;
                default:
                    resolve(strFieldValue);
            }
        });
    }

    /**
     * Gets property of the Field by Field's ID and Property Name
     * @param fieldId Field's ID
     * @param propertyName Property name
     * @param context SPFx context
     */
    public static getFieldProperty(fieldId: string, propertyName: string, context: IContext): Promise<any> {
        return new Promise<any>(resolve => {
            let loadedViewFields: { [viewId: string]: IFields } = SPHelper._getLoadedViewFieldsFromStorage();
            const viewId: string = SPHelper.getPageViewId(context);

            if (!loadedViewFields) {
                loadedViewFields = {};
            }

            if (!loadedViewFields[viewId]) {
                loadedViewFields[viewId] = {};
            }

            let field: ISPField = loadedViewFields[viewId][fieldId];
            if (!field) {
                field = {
                    Id: fieldId
                };
            }


            if (GeneralHelper.isDefined(field[propertyName])) {
                resolve(field[propertyName]);
                return;
            }

            sp.setup({
                spfxContext: context
            });

            sp.web.lists.getByTitle(context.pageContext.list.title).fields.getById(fieldId).select(propertyName).get().then(f => {
                field[propertyName] = f[propertyName];

                loadedViewFields[viewId][field.Id] = field;

                SPHelper._updateSessionStorageLoadedViewFields(loadedViewFields);
                resolve(field);
            });
        });
    }

    /**
     * Gets column's value for the row using List Item Accessor.
     * This method works with private property _values of List Item Accessor to get such values as FriendlyDisplay text for Date, and more.
     * @param listItem List Item Accessor
     * @param itemName column name
     */
    public static getRowItemValueByName(listItem: ListItemAccessor, itemName: string): any {
        const _values: any = (<any>listItem)._values;

        if (_values) {
            return (_values as Map<string, any>).get(itemName);
        }
        else {
            //
            // TODO: here we should call make a POST request to _api/web/GetList(@listUrl)/RenderListDataAsStream with correct parameters to get correct data
            // the parameters should contain view, folder, pagination data, etc.
            // I hope that Dev team will expose this data in API before I implement that because it's pretty complicated and they already have it in place
            //

            return null;
        }
    }

    /**
     * Gets SchemaXml for the field by List Title and Field Id
     * @param fieldId Field's Id
     * @param listTitle List Title
     * @param context Customizer's context
     */
    public static getFieldSchemaXmlById(fieldId: string, listTitle: string, context: IContext): Promise<string> {
        return new Promise<string>((resolve) => {
            let loadedViewFields: { [viewId: string]: IFields } = SPHelper._getLoadedViewFieldsFromStorage();
            const viewId: string = SPHelper.getPageViewId(context);

            if (!loadedViewFields) {
                loadedViewFields = {};
            }

            if (!loadedViewFields[viewId]) {
                loadedViewFields[viewId] = {};
            }

            let field: ISPField = loadedViewFields[viewId][fieldId];
            if (!field) {
                field = {
                    Id: fieldId
                };
            }


            if (GeneralHelper.isDefined(field.SchemaXml)) {
                resolve(field.SchemaXml);
                return;
            }
            sp.setup({
                spfxContext: context
            });

            sp.web.lists.getByTitle(listTitle).fields.getById(fieldId).select('SchemaXml').get().then((f) => {
                field.SchemaXml = f && f.SchemaXml;

                loadedViewFields[viewId][field.Id] = field;

                SPHelper._updateSessionStorageLoadedViewFields(loadedViewFields);
                resolve(f ? f.SchemaXml : '');
            });
        });
    }

    /**
     * Gets correct view id from the page
     * @param context SPFx Context
     */
    public static getPageViewId(context: IContext): string {
        const urlParams: URLSearchParams = new URLSearchParams(location.search);
        let viewIdQueryParam: string = urlParams.get('viewid');
        if (viewIdQueryParam && viewIdQueryParam.indexOf('{') !== 0) {
            viewIdQueryParam = `{${viewIdQueryParam}}`;
        }
        return viewIdQueryParam || context.pageContext.legacyPageContext.viewId;
    }

    private static _updateSessionStorageLoadedViewFields(loadedViewFields: { [viewId: string]: IFields }): void {
        const sessionStorage: any = window.sessionStorage;
        sessionStorage.setItem(Constants.LoadedViewFieldsKey, JSON.stringify(loadedViewFields));
    }

    private static _getLoadedViewFieldsFromStorage(): { [viewId: string]: IFields } {
        const loadedViewFields = sessionStorage.getItem(Constants.LoadedViewFieldsKey);
        if (loadedViewFields) {
            return JSON.parse(loadedViewFields);
        }

        return null;
    }
}