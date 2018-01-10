import { SPHttpClient } from '@microsoft/sp-http';
import { PageContext, SPField } from '@microsoft/sp-page-context';
import { ListViewAccessor } from "@microsoft/sp-listview-extensibility";

/**
 * Customizer context interface.
 * Can be used in different types of customizers
 */
export interface IContext {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    listView?: ListViewAccessor | null;
    field?: SPField | null;
}

/**
 * Custom interface for web regional settings.
 * It contains some additional information like timezone offset and date format
 */
export interface IRegionalSettings {
    thousandSeparator: string;
    decimalSeparator: string;
    hoursOffset: number;
    webDateFormat: string;
    workDays: number;
    firstDayOfWeek: number;
    firstWeekOfYear: number;
}

/**
 * Parent of all props interfaces that needs context
 */
export interface IProps {
    context: IContext;
}

/**
 * Digit separators interface
 */
export interface IDigitSeparators {
    thousand: string;
    decimal: string;
}