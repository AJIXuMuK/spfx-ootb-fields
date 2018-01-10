import { IContext, IRegionalSettings, IDigitSeparators } from '../common/Interfaces';
import { SPHelper } from './SPHelper';
import '../common/Extensions/String.extensions';

import * as _ from 'lodash';

import * as strings from 'OotbFieldsFieldCustomizerStrings';

var moment: any = require('moment');

/**
 * Helper with general methods to simplify some routines
 */
export class GeneralHelper {
    /**
     * Trims slash at the end of URL if needed
     * @param url URL
     */
    public static trimSlash(url: string): string {
        if (url.lastIndexOf('/') === url.length - 1)
            return url.slice(0, -1);
        return url;
    }

    /**
     * Loads predefined JSOM scripts in correct sequence
     * @param context Customizer's context
     */
    public static loadJSOMScriptsInSequence(context: IContext): Promise<void> {
        return new Promise<void>((resolve) => {
            this.loadScripts(context, ['MicrosoftAjax.js'])
                .then(() => {
                    return this.loadScripts(context, ['init.js']);
                })
                .then(() => {
                    return this.loadScripts(context, ['sp.runtime.js']);
                })
                .then(() => {
                    return this.loadScripts(context, ['sp.js']);
                })
                .then(() => {
                    resolve();
                });
        });
    }

    /**
     * Loads neede scripts
     * @param context Customizer's context
     * @param scripts Scripts to load
     * @param fromLayouts flag if scripts should be loaded from /_layouts/ folder
     */
    public static loadScripts(context: IContext, scripts: Array<string>, fromLayouts?: boolean): Promise<void> {
        return new Promise<void>((resolve) => {
            fromLayouts = fromLayouts !== false;

            const promises: Array<Promise<void>> = scripts.map((script): Promise<void> => {
                let fullSrc: string = fromLayouts ? GeneralHelper.trimSlash(context.pageContext.site.absoluteUrl) + '/_layouts/15/' + script : script;
                return GeneralHelper.loadScript(fullSrc);
            });

            Promise.all(promises).then(() => {
                resolve();
            });
        });
    }

    /**
     * Loads script
     * @param url: script src
     * @param globalObjectName: name of global object to check if it is loaded to the page
     */
    public static loadScript(url: string, globalObjectName?: string): Promise<void> {
        return new Promise<void>((resolve) => {
            let isLoaded = false;
            if (globalObjectName) {
                isLoaded = false;
                if (globalObjectName.indexOf('.') !== -1) {
                    const props = globalObjectName.split('.');
                    let currObj: any = window;

                    for (let i = 0, len = props.length; i < len; i++) {
                        if (!currObj[props[i]]) {
                            isLoaded = false;
                            break;
                        }

                        currObj = currObj[props[i]];
                    }
                }
                else {
                    isLoaded = !!window[globalObjectName];
                }
            }
            // checking if the script was previously added to the page
            if (isLoaded || document.head.querySelector('script[src="' + url + '"]')) {
                resolve();
                return;
            }

            // loading the script
            const script = document.createElement('script');
            script.type = 'text/javascript';
            script.src = url;
            script.onload = () => {
                resolve();
            };
            document.head.appendChild(script);
        });
    }

    /**
     * Checks if the value is a numeric value
     * @param value value
     */
    public static isNumeric(value: any): boolean {
        const type = typeof value;

        return (type === 'number' || type === 'string') && !isNaN(value - parseFloat(value));
    }

    /**
     * Checks if the date is a valid date
     * @param date date 
     */
    public static isValidDate(date: any): boolean {
        let val: any;
        if (typeof date === 'string')
            val = new Date(date);
        else
            val = date;

        return (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val.getTime()));
    }

    /**
     * Gets Regular expression to test if value is a correct digit in current web locale
     * @param context SPFx context
     */
    public static getDigitTestRegExp(): RegExp {
        const regSettings: IRegionalSettings = SPHelper.getWebRegionalSettings();
        return RegExp('(^\\d+(' + regSettings.decimalSeparator + '\\d+)?$)|(^\\d{1,3}('
            + regSettings.thousandSeparator + '\\d{3})+(' + regSettings.decimalSeparator + '\\d+)?$)', 'gmi');
    }

    /**
     * Formats digit in current culture
     * @param digit digit string in invariant culture
     */
    public static formatDigit(digit: string): string {
        if (!digit)
            return;
        var thousandRegExp = /(\d+)(\d{3})/,
            digitParts = digit.split('.'),
            digitIntPart = digitParts[0],
            digitDecPart = digitParts[1],
            seps = this.getDigitSeparators();

        while (thousandRegExp.test(digitIntPart)) {
            digitIntPart = digitIntPart.replace(thousandRegExp, '$1' + seps.thousand + '$2');
        }

        if (digitDecPart)
            return digitIntPart + seps.decimal + digitDecPart;

        return digitIntPart;
    }

    /**
     * Gets digit separators
     */
    public static getDigitSeparators(): IDigitSeparators {
        const separators: IDigitSeparators = {
            thousand: ',',
            decimal: '.'
        };

        const regSettings: IRegionalSettings = SPHelper.getWebRegionalSettings();
        if (regSettings) {
            separators.thousand = regSettings.thousandSeparator;
            separators.decimal = regSettings.decimalSeparator;
        }

        return separators;
    }
    /**
     * Checks if the digit string is a correct digit string in current locale
     * @param digit digit string to test
     */
    public static isCorrectLocaleDigit(digit: string): boolean {
        const digitRegExp: RegExp = this.getDigitTestRegExp();

        return !digit || digitRegExp.test(digit.trim());
    }

    /**
     * 
     * @param date formats date based on Web Regional Settings
     */
    public static formatDate(date: Date): string {
        const regSettings: IRegionalSettings = SPHelper.getWebRegionalSettings();
        if (regSettings) {
            return moment(date).format(regSettings.webDateFormat);
        }
        return moment(date).format('MM/DD/YYYY');
    }

    /**
     * De-formats date string from Web Regional Settings format to MM/DD/YYYY string representation
     * @param date date string
     */
    public static deFormatDate(date: string): string {
        const regSettings: IRegionalSettings = SPHelper.getWebRegionalSettings();
        if (regSettings) {
            return moment(date, regSettings.webDateFormat, true).format('MM/DD/YYYY');
        }
        return date;
    }

    /**
     * Parses date string based on current Web Regional Settings
     * @param date date string
     */
    public static parseDateCurrentRegionalSettings(date: string): Date {
        const regSettings: IRegionalSettings = SPHelper.getWebRegionalSettings();
        if (regSettings) {
            return moment(date, regSettings.webDateFormat);
        }
        return new Date(date);
    }

    /**
     * De-formats digit string from Web Regional Settings locale string to invariant culture digit string
     * @param digit digit string in Web Regional Settings local
     */
    public static deFormatDigit(digit: string): string {
        const seps: IDigitSeparators = this.getDigitSeparators();

        const thousandSep: string = this.escapeRegExpSpecialCharacter(seps.thousand);
        const decimalSep: string = this.escapeRegExpSpecialCharacter(seps.decimal);

        return digit.trim().replace(new RegExp(thousandSep, 'gmi'), '').replace(new RegExp(decimalSep, 'gmi'), '.');
    }

    /**
     * Escapes \s and \. characters
     * @param char character
     */
    public static escapeRegExpSpecialCharacter(char: string): string {
        if (char.length === 1) {
            if (/\s/.test(char))
                return '\\s';
            if (/\./.test(char))
                return '\\.';
        }

        return char;
    }

    /**
     * Encodes text
     * @param text text to encode
     */
    public static encodeText(text: string): string {
        const n = /[<>&'"\\]/g;
        return text ? text.replace(n, this._getEncodedChar) : '';
    }

    /**
     * Copy of Microsoft's GetRelativeDateTimeString from SP.dateTimeUtil
     */
    public static getRelativeDateTimeString(format: string): string {
        const formatParts: string[] = format.split('|');
        let result: string = null;
        let placeholdersString: string = null;

        if (formatParts[0] == '0')
            return format.substring(2);
        const isFuture: boolean = formatParts[1] === '1';
        const formatType: string = formatParts[2];
        const timeString: string = formatParts.length >= 4 ? formatParts[3] : null;
        const dayString: string = formatParts.length >= 5 ? formatParts[4] : null;

        switch (formatType) {
            case '1':
                result = isFuture ? strings.DateTime['L_RelativeDateTime_AFewSecondsFuture'] : strings.DateTime['L_RelativeDateTime_AFewSeconds'];
                break;
            case '2':
                result = isFuture ? strings.DateTime['L_RelativeDateTime_AboutAMinuteFuture'] : strings.DateTime['L_RelativeDateTime_AboutAMinute'];
                break;
            case '3':
                placeholdersString = this.getLocalizedCountValue(isFuture ? strings.DateTime['L_RelativeDateTime_XMinutesFuture'] : strings.DateTime['L_RelativeDateTime_XMinutes'], isFuture ? strings.DateTime['L_RelativeDateTime_XMinutesFutureIntervals'] : strings.DateTime['L_RelativeDateTime_XMinutesIntervals'], Number(timeString));
                break;
            case '4':
                result = isFuture ? strings.DateTime['L_RelativeDateTime_AboutAnHourFuture'] : strings.DateTime['L_RelativeDateTime_AboutAnHour'];
                break;
            case '5':
                if (timeString == null) {
                    result = isFuture ? strings.DateTime['L_RelativeDateTime_Tomorrow'] : strings.DateTime['L_RelativeDateTime_Yesterday'];
                }
                else {
                    placeholdersString = isFuture ? strings.DateTime['L_RelativeDateTime_TomorrowAndTime'] : strings.DateTime['L_RelativeDateTime_YesterdayAndTime'];
                }
                break;
            case '6':
                placeholdersString = this.getLocalizedCountValue(
                    isFuture ? strings.DateTime['L_RelativeDateTime_XHoursFuture'] : strings.DateTime['L_RelativeDateTime_XHours'],
                    isFuture ? strings.DateTime['L_RelativeDateTime_XHoursFutureIntervals'] : strings.DateTime['L_RelativeDateTime_XHoursIntervals'],
                    Number(timeString));
                break;
            case '7':
                if (dayString == null) {
                    result = timeString;
                }
                else {
                    placeholdersString = strings.DateTime['L_RelativeDateTime_DayAndTime'];
                }
                break;
            case '8':
                placeholdersString = this.getLocalizedCountValue(
                    isFuture ? strings.DateTime['L_RelativeDateTime_XDaysFuture'] : strings.DateTime['L_RelativeDateTime_XDays'],
                    isFuture ? strings.DateTime['L_RelativeDateTime_XDaysFutureIntervals'] : strings.DateTime['L_RelativeDateTime_XDaysIntervals'],
                    Number(timeString));
                break;
            case '9':
                result = strings.DateTime['L_RelativeDateTime_Today'];
        }
        if (placeholdersString != null) {
            result = placeholdersString.replace("{0}", timeString);
            if (dayString != null) {
                result = result.replace("{1}", dayString);
            }
        }
        return result;
    }

    /**
     * Copy of Microsoft's GetLocalizedCountValue from SP.dateTimeUtil.
     * I've tried to rename all the vars to have meaningful names... but some were too unclear
     */
    public static getLocalizedCountValue(format: string, first: string, second: number): string {
        if (format == undefined || first == undefined || second == undefined)
            return null;
        let result: string = '';
        let a = -1;
        let firstOperandOptions: string[] = first.split('||');

        for (let firstOperandOptionsIdx = 0, firstOperandOptionsLen = firstOperandOptions.length; firstOperandOptionsIdx < firstOperandOptionsLen; firstOperandOptionsIdx++) {
            const firstOperandOption: string = firstOperandOptions[firstOperandOptionsIdx];

            if (firstOperandOption == null || firstOperandOption === '')
                continue;
            let optionParts: string[] = firstOperandOption.split(',');

            for (var optionPartsIdx = 0, optionPartsLen = optionParts.length; optionPartsIdx < optionPartsLen; optionPartsIdx++) {
                const optionPart: string = optionParts[optionPartsIdx];

                if (optionPart == null || optionPart === '')
                    continue;
                if (isNaN(optionPart.parseNumberInvariant())) {
                    const dashParts: string[] = optionPart.split('-');

                    if (dashParts == null || dashParts.length !== 2)
                        continue;
                    var j, n;

                    if (dashParts[0] === '')
                        j = 0;
                    else if (isNaN(dashParts[0].parseNumberInvariant()))
                        continue;
                    else
                        j = parseInt(dashParts[0]);
                    if (second >= j) {
                        if (dashParts[1] === '') {
                            a = firstOperandOptionsIdx;
                            break;
                        }
                        else if (isNaN(dashParts[1].parseNumberInvariant()))
                            continue;
                        else
                            n = parseInt(dashParts[1]);
                        if (second <= n) {
                            a = firstOperandOptionsIdx;
                            break;
                        }
                    }
                }
                else {
                    var p = parseInt(optionPart);

                    if (second === p) {
                        a = firstOperandOptionsIdx;
                        break;
                    }
                }
            }
            if (a !== -1)
                break;
        }
        if (a !== -1) {
            var e = format.split('||');

            if (e != null && e[a] != null && e[a] != '')
                result = e[a];
        }
        return result;
    }

    /**
     * Extracts text from HTML strings without creating HTML elements
     * @param html HTML string
     */
    public static getTextFromHTML(html: string): string {
        let result: string = html;
        let oldResult: string = result;
        const tagBody = '(?:[^"\'>]|"[^"]*"|\'[^\']*\')*';

        const tagOrComment = new RegExp(
            '<(?:'
            // Comment body.
            + '!--(?:(?:-*[^->])*--+|-?)'
            // Special "raw text" elements whose content should be elided.
            + '|script\\b' + tagBody + '>[\\s\\S]*?</script\\s*'
            + '|style\\b' + tagBody + '>[\\s\\S]*?</style\\s*'
            // Regular name
            + '|/?[a-z]'
            + tagBody
            + ')>',
            'gi');

        do {
            oldResult = result;
            result = result.replace(tagOrComment, '');
        } while (result !== result);

        return result;
    }

    private static _getEncodedChar(c): string {
        const o = {
            "<": "&lt;",
            ">": "&gt;",
            "&": "&amp;",
            '"': "&quot;",
            "'": "&#39;",
            "\\": "&#92;"
        };
        return o[c];
    }
}