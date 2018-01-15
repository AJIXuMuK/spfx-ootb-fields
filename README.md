# spfx-ootb-fields

This repository contains a set of React components that can be used in SPFx Field Customizers to provide rendering of the fields similar to out of the box experience. Additional benefit is ability to set custom css classes and styles to the component.
Related UserVoice requests:[https://sharepoint.uservoice.com/forums/329220-sharepoint-dev-platform/suggestions/18810637-access-to-re-use-modern-field-render-controls](https://sharepoint.uservoice.com/forums/329220-sharepoint-dev-platform/suggestions/18810637-access-to-re-use-modern-field-render-controls)[https://sharepoint.uservoice.com/forums/329220-sharepoint-dev-platform/suggestions/31530607-field-customizer-ability-to-call-ootb-render-meth](https://sharepoint.uservoice.com/forums/329220-sharepoint-dev-platform/suggestions/31530607-field-customizer-ability-to-call-ootb-render-meth)

## Getting started
### Installation
To get started you need to install this package to your project and also dependency package `@microsoft/sp-dialog`.

Enter the following commands to install dependencies to your project:
```
npm i spfx-ootb-fields --save
npm i @microsoft/sp-dialog
```

### Configuration
Once the package is installed, you will have to configure the resource file of the property controls to be used in your project. You can do this by opening the `config/config.json` and adding the following line to the `localizedResources` property:
```
"OotbFieldsStrings": "./node_modules/spfx-ootb-fields/lib/loc/{locale}.js"
```
Also, you will have to the url to load `moment.js` library from the CDN. It is done in `config/config.json` file as well by adding next line to the `externals` section:
```
"moment": "https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.20.1/moment.min.js"
```

## Usage
The main scenario to use this package is to import `FieldRendererHelper` class and to call `getFieldRenderer` method. This method returns proper field renderer (`JSX.Element`) based on field's type. It means that it will automatically select proper component that should be rendered in this or that field. This method also contains logic to correctly process field's value and get correct text to display (for example, Friendly Text for DateTime fields).
Here is an example on how it can be used inside custom Field Customizer component (.tsx file):
```
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
```

Additionally, any of icluded components can be used by itself.

## Controls
Here is a list of the controls included in the package with a description which field types are covered with the specific control.
All controls contain next common properties in React props object:
`cssProps?: React.CSSProperties` - CSS properties to apply to the renderer
`className?: ICssInput` - CSS classes to apply to the renderer

| Control | Props | Fields Covered | Comments |
| --- | --- | --- | --- |
| AttachmentsRenderer | `count?: number` - amount of attachments | Attachments | Renders Clip icon if there are attachments for the current list item |
| DateRenderer | `text?: string` - text to be rendered | Date and Time | Renders date and time value as simple text |
| FileTypeRenderer | `path: string` - document path<br>`isFolder?: boolean` - true if the icon should be rendered for a folder, not file | DocIcon | Renders an icon based on the extension of the current document. Office UI Fabric icons font is used to render the icons |
| LookupRenderer | `lookups: ISPFieldLookupValue[]` - lookup values<br>`dispFormUrl?: string` - url of Display form for the list that is referenced in the lookup<br>`onClick?: (args: ILookupClickEventArgs) => {}` - custom event handler of lookup item click. If not set the dialog with Display Form will be shown | Lookup (single and multi) | Renders each referenced value as a link on a separate line. Opens popup with Display Form when the link is clicked |
| NameRenderer | `text?: string` - text to display<br>`isLink: boolean` - if the Name should be rendered as link<br>`filePath?: string` - path to the document<br>`isNew?: boolean` - true if the document is new<br>`hasPreview?: boolean` - true if the document type has preview (true by default)<br>`onClick?: (args: INameClickEventArgs) => {}` - custom handler for link click. If not set link click will lead to rendering document preview | Document's Name (FileLeafRef, LinkFilename, LinkFilenameNoMenu) | Renders document's name as a link. The link provides either preview (if it is available) or direct download. Additionally, new documents are marked with "Glimmer" icon |
| TaxonomyRenderer | `terms: ITerm[]` - terms to display | Managed Metadata | Renders each term on a separate line |
| TextRenderer | `text?: string` - text to display<br>`isSafeForInnerHTML?: boolean` - true if props.text can be inserted as innerHTML of the component<br>`isTruncated?: boolean` - true if the text should be truncated | Single line of text; Multiple lines of text; Choice (single and multi); Yes/No; Integer; Counter; Number; Currency; also used as a default renderer for not implemented fields | Renders field's value as a simple text or HTML |
| TitleRenderer | `text?: string` - text to display<br>`isLink?: boolean` - true if the Title should be rendered as link<br>`baseUrl?: string` - web url<br>`listId?: string` - list id<br>`id?: number` - item id<br>`onClick?: (args: ITitleClickEventArgs) => {}` - custom title click event handler. If not set Display form for the item will be displayed | List Item's Title (Title, LinkTitleNoMenu, LinkTitle) | The control renders title either as a simple text or as a link on the Dislpay Form. Additional actions like Share and Context Menu are not implemented |
| UrlRenderer | `text?: string` - text to display<br>`url?: string` - url<br>`isImageUrl?: boolean` - true if the field should be rendered as an image | Hyperlink or Image; URL field from Links List | Renders either link or image |
| UserRenderer | `users?: IPrincipal[]` - users/groups to be displayed<br>`context: IContext` - customizer's context | People and Group | Renders each referenced user/group as a link on a separate line. Hovering the link for users (not groups) leads to opening of Persona control. |

## Utilities Classes
Here is a list of Utilities classes and public methods that are included in the package and could be also helpful:
<table>
<thead>
<tr>
<th>Class</th><th>Method</th><th>Description</th>
</tr>
</thead>
<tbody>
<tr>
<td>`FieldRenderer`</td>
<td>`getFieldRenderer(fieldValue: any, props: IFieldRendererProps, listItem: ListItemAccessor, context: IContext): JSX.Element`</td>
<td>Returns JSX.Element with OOTB rendering and applied additional props.<br>
Parameters<br>
`fieldValue` - Value of the field<br>
`props` - IFieldRendererProps (CSS classes and CSS styles)<br>
`listItem` - Current list item<br>
`context` - Customizer's context
</td>
</tr>
<tr>
<td rowspan="19">
`GeneralHelper`
</td>
<td>
`trimSlash(url: string): string`
</td>
<td>
Trims slash at the end of URL if needed<br>
Parameters<br>
`url` - URL
</td>
</tr>
<tr>
<td>
`loadJSOMScriptsInSequence(context: IContext): Promise<void>`
</td>
<td>
Loads predefined JSOM scripts in correct sequence<br>
Parameters<br>
`context` - Customizer's context
</td>
</tr>
<tr>
<td>
`loadScripts(context: IContext, scripts: Array<string>, fromLayouts?: boolean): Promise<void>`
</td>
<td>
Loads needed scripts<br>
Parameters<br>
`context` - Customizer's context<br>
`scripts` - Scripts to load<br>
`fromLayouts` - flag if scripts should be loaded from /_layouts/15 folder
</td>
</tr>
<tr>
<td>
`loadScript(url: string, globalObjectName?: string): Promise<void>`
</td>
<td>
Loads script<br>
Parameters<br>
`url` - script src<br>
`globalObjectName` - name of global object to check if it is loaded to the page
</td>
</tr>
<tr>
<td>
`isNumeric(value: any): boolean`
</td>
<td>
Checks if the value is a numeric value<br>
Parameters</br>
`value` - value
</td>
</tr>
<tr>
<td>
`isValidDate(date: any): boolean`
</td>
<td>
Checks if the date is a valid date<br>
Parameters<br>
`date` - date
</td>
</tr>
<tr>
<td>
`getDigitTestRegExp(): RegExp`
</td>
<td>
Gets Regular expression to test if value is a correct digit in current web locale
</td>
</tr>
<tr>
<td>
`formatDigit(digit: string): string`
</td>
<td>
Formats digit in current culture<br>
Parameters<br>
`digit` - digit string in invariant culture
</td>
</tr>
<tr>
<td>
`getDigitSeparators(): IDigitSeparators`
</td>
<td>
Gets digit separators
</td>
</tr>
<tr>
<td>
`isCorrectLocaleDigit(digit: string): boolean`
</td>
<td>
Checks if the digit string is a correct digit string in current locale<br>
Parameters<br>
`digit` - digit string to test
</td>
</tr>
<tr>
<td>
`formatDate(date: Date): string`
</td>
<td>
formats date based on Web Regional Settings<br>
Parameters<br>
`date` - date
</td>
</tr>
<tr>
<td>
`deFormatDate(date: string): string`
</td>
<td>
De-formats date string from Web Regional Settings format to MM/DD/YYYY string representation<br>
Parameters<br>
`date` - date string
</td>
</tr>
<tr>
<td>
`parseDateCurrentRegionalSettings(date: string): Date`
</td>
<td>
Parses date string based on current Web Regional Settings<br>
Parameters<br>
`date` - date string
</td>
</tr>
<tr>
<td>
`deFormatDigit(digit: string): string`
</td>
<td>
De-formats digit string from Web Regional Settings locale string to invariant culture digit string<br>
Parameters<br>
`digit` - digit string in Web Regional Settings locale
</td>
</tr>
<tr>
<td>
`escapeRegExpSpecialCharacter(char: string): string`
</td>
<td>
Escapes \s and \. characters<br>
Parameters<br>
`char` - character
</td>
</tr>
<tr>
<td>
`encodeText(text: string): string`
</td>
<td>
Encodes text<br>
Parameters<br>
`text` - text to encode
</td>
</tr>
<tr>
<td>
`getRelativeDateTimeString(format: string): string`
</td>
<td>
Copy of Microsoft's GetRelativeDateTimeString from SP.dateTimeUtil
</td>
</tr>
<tr>
<td>
`getLocalizedCountValue(format: string, first: string, second: number): string`
</td>
<td>
Extracts text from HTML strings without creating HTML elements<br>
Parameters<br>
`html` - HTML string
</td>
</tr>
</tbody>
</table>

### Debug Url
?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&fieldCustomizers={"Percent":{"id":"57ebd944-98ed-43f9-b722-e959d6dac6ad"}}

