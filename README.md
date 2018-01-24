# spfx-ootb-fields

This repository contains a set of React components that can be used in SPFx Field Customizers to provide rendering of the fields similar to out of the box experience. Additional benefit is ability to set custom css classes and styles to the component.
Related UserVoice requests:<br>
[https://sharepoint.uservoice.com/forums/329220-sharepoint-dev-platform/suggestions/18810637-access-to-re-use-modern-field-render-controls](https://sharepoint.uservoice.com/forums/329220-sharepoint-dev-platform/suggestions/18810637-access-to-re-use-modern-field-render-controls)<br>
[https://sharepoint.uservoice.com/forums/329220-sharepoint-dev-platform/suggestions/31530607-field-customizer-ability-to-call-ootb-render-meth](https://sharepoint.uservoice.com/forums/329220-sharepoint-dev-platform/suggestions/31530607-field-customizer-ability-to-call-ootb-render-meth)

## Getting started
### Installation
To get started you need to install this package to your project and also dependency package `@microsoft/sp-dialog`.

Enter the following commands to install dependencies to your project:
```
npm i spfx-ootb-fields --save
npm i @microsoft/sp-dialog --save
```

### Configuration
Once the package is installed, you will have to configure the resource file of the property controls to be used in your project. You can do this by opening the `config/config.json` and adding the following line to the `localizedResources` property:
```
"OotbFieldsStrings": "./node_modules/spfx-ootb-fields/lib/loc/{locale}.js"
```

## Usage
The main scenario to use this package is to import `FieldRendererHelper` class and to call `getFieldRenderer` method. This method returns a Promise with a proper field renderer (`Promise<JSX.Element>`) based on field's type. It means that it will automatically select proper component that should be rendered in this or that field. This method also contains logic to correctly process field's value and get correct text to display (for example, Friendly Text for DateTime fields).
As the method returns `Promise` it should be called in one of React component lifecycle methods, for example, `componentWillMount` that will occur before `render`. The resulting field renderer could be saved in the element's state and used later in `render` method.
Here is an example on how it can be used inside custom Field Customizer component (.tsx file):
```
export interface IOotbFieldsState {
  fieldRenderer?: JSX.Element;
}

//...

@override
  public componentWillMount() {
    FieldRendererHelper.getFieldRenderer(this.props.value, {
      className: this.props.className,
      cssProps: this.props.cssProps
    }, this.props.listItem, this.props.context).then(fieldRenderer => {
      this.setState({
        fieldRenderer: fieldRenderer
      });
    });
  }

public render(): React.ReactElement<{}> {
    return (
      <div className={styles.cell}>
        {this.state.fieldRenderer}
      </div>
    );
  }
```

Additionally, any of included components can be used by itself.

## Controls
Here is a list of the controls included in the package with a description which field types are covered with the specific control.
All controls contain next common properties in React props object:
`cssProps?: React.CSSProperties` - CSS properties to apply to the renderer
`className?: ICssInput` - CSS classes to apply to the renderer

| Control | Props | Fields Covered | Comments |
| --- | --- | --- | --- |
| FieldAttachmentsRenderer | `count?: number` - amount of attachments | Attachments | Renders Clip icon if there are attachments for the current list item |
| FieldDateRenderer | `text?: string` - text to be rendered | Date and Time | Renders date and time value as simple text |
| FieldFileTypeRenderer | `path: string` - document path<br>`isFolder?: boolean` - true if the icon should be rendered for a folder, not file | DocIcon | Renders an icon based on the extension of the current document. Office UI Fabric icons font is used to render the icons |
| FieldLookupRenderer | `lookups: ISPFieldLookupValue[]` - lookup values<br>`dispFormUrl?: string` - url of Display form for the list that is referenced in the lookup<br>`onClick?: (args: ILookupClickEventArgs) => {}` - custom event handler of lookup item click. If not set the dialog with Display Form will be shown | Lookup (single and multi) | Renders each referenced value as a link on a separate line. Opens popup with Display Form when the link is clicked |
| FieldNameRenderer | `text?: string` - text to display<br>`isLink: boolean` - if the Name should be rendered as link<br>`filePath?: string` - path to the document<br>`isNew?: boolean` - true if the document is new<br>`hasPreview?: boolean` - true if the document type has preview (true by default)<br>`onClick?: (args: INameClickEventArgs) => {}` - custom handler for link click. If not set link click will lead to rendering document preview | Document's Name (FileLeafRef, LinkFilename, LinkFilenameNoMenu) | Renders document's name as a link. The link provides either preview (if it is available) or direct download. Additionally, new documents are marked with "Glimmer" icon |
| FieldTaxonomyRenderer | `terms: ITerm[]` - terms to display | Managed Metadata | Renders each term on a separate line |
| FieldTextRenderer | `text?: string` - text to display<br>`isSafeForInnerHTML?: boolean` - true if props.text can be inserted as innerHTML of the component<br>`isTruncated?: boolean` - true if the text should be truncated | Single line of text; Multiple lines of text; Choice (single and multi); Yes/No; Integer; Counter; Number; Currency; also used as a default renderer for not implemented fields | Renders field's value as a simple text or HTML |
| FieldTitleRenderer | `text?: string` - text to display<br>`isLink?: boolean` - true if the Title should be rendered as link<br>`baseUrl?: string` - web url<br>`listId?: string` - list id<br>`id?: number` - item id<br>`onClick?: (args: ITitleClickEventArgs) => {}` - custom title click event handler. If not set Display form for the item will be displayed | List Item's Title (Title, LinkTitleNoMenu, LinkTitle) | The control renders title either as a simple text or as a link on the Dislpay Form. Additional actions like Share and Context Menu are not implemented |
| FieldUrlRenderer | `text?: string` - text to display<br>`url?: string` - url<br>`isImageUrl?: boolean` - true if the field should be rendered as an image | Hyperlink or Image; URL field from Links List | Renders either link or image |
| FieldUserRenderer | `users?: IPrincipal[]` - users/groups to be displayed<br>`context: IContext` - customizer's context | People and Group | Renders each referenced user/group as a link on a separate line. Hovering the link for users (not groups) leads to opening of Persona control. |

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
<td><code>FieldRenderer</code></td>
<td><code>getFieldRenderer(fieldValue: any, props: IFieldRendererProps, listItem: ListItemAccessor, context: IContext): Promise&lt;JSX.Element&gt;</code></td>
<td>Returns <code>Promise&lt;JSX.Element&gt;</code> with OOTB rendering and applied additional props.<br>
Parameters<br>
<code>fieldValue</code> - Value of the field<br>
<code>props</code> - IFieldRendererProps (CSS classes and CSS styles)<br>
<code>listItem</code> - Current list item<br>
<code>context</code> - Customizer's context
</td>
</tr>
<tr>
<td rowspan="7">
<code>GeneralHelper</code>
</td>
<td>
<code>trimSlash(url: string): string</code>
</td>
<td>
Trims slash at the end of URL if needed<br>
Parameters<br>
<code>url</code> - URL
</td>
</tr>
<tr>
<td>
<code>encodeText(text: string): string</code>
</td>
<td>
Encodes text<br>
Parameters<br>
<code>text</code> - text to encode
</td>
</tr>
<tr>
<td>
<code>getRelativeDateTimeString(format: string): string</code>
</td>
<td>
Copy of Microsoft's GetRelativeDateTimeString from SP.dateTimeUtil
</td>
</tr>
<tr>
<td>
<code>getLocalizedCountValue(format: string, first: string, second: number): string</code>
</td>
<td>
Copy of Microsoft's GetLocalizedCountValue from SP.dateTimeUtil.<br>
I've tried to rename all the vars to have meaningful names... but some were too unclear
</td>
</tr>
<tr>
<td>
<code>getTextFromHTML(html: string): string</code>
</td>
<td>
Extracts text from HTML strings without creating HTML elements<br>
Parameters<br>
<code>html</code> - HTML string
</td>
</tr>
<tr>
<td>
<code>isDefined(value: any): boolean</code>
</td>
<td>
Checks if value is defined (not null and not undefined)<br>
<code>value</code> - value
</td>
</tr>
<tr>
<td>
<code>parseXml(xmlString: string): Document</code>
</td>
<td>
Creates Document element based on Xml string<br>
<code>xmlString</code> - XML string to parse
</td>
</tr>

<tr>
<td rowspan="7">
<code>SPHelper</code>
</td>
<td>
<code>getStoredFieldName(columnName: string): string</code>
</td>
<td>
Gets field's Real Name from FieldNamesMapping<br>
Parameters<br>
<code>columnName</code> - current field name
</td>
</tr>
<tr>
<td>
<code>getFieldText(fieldValue: any, listItem: ListItemAccessor, context: IContext): Promise&lt;string&gt;</code>
</td>
<td>
Asynchronously gets Field's text<br>
Parameters<br>
<code>fieldValue</code> - field value as it appears in Field Customizer's onRenderCell event
<code>listItem</code> - List Item accessor
<code>context</code> - Customizer's context
</td>
</tr>
</tr>
<tr>
<td>
<code>getFieldProperty(fieldId: string, propertyName: string, context: IContext, fromSchemaXml: boolean): Promise&lt;any&gt;</code>
</td>
<td>
Asynchronously gets property of the Field by Field's ID and Property Name<br>
Parameters<br>
<code>fieldId</code> - Field's ID<br>
<code>propertyName</code> - Property name<br>
<code>context</code> - SPFx context<br>
<code>fromSchemaXml</code> - true if the field should be read from Field Schema Xml
</td>
</tr>
<tr>
<td>
<code>getLookupFieldListDispFormUrl(fieldId: string, context: IContext): Promise&lt;any&gt;</code>
</td>
<td>
Asynchronously gets the Diplay Form Url for the Lookup field<br>
Parameters<br>
<code>fieldId</code> - Field's ID<br>
<code>context</code> - SPFx context<br>
</td>
</tr>
<tr>
<td>
<code>getRowItemValueByName(listItem: ListItemAccessor, itemName: string): any</code>
</td>
<td>
Gets column's value for the row using List Item Accessor.<br>
This method works with private property <code>_values</code> of List Item Accessor to get such values as FriendlyDisplay text for Date, and more.<br>
Parameters<br>
<code>listItem</code> - List Item Accessor<br>
<code>itemName<code> - column name
</td>
</tr>
<tr>
<td>
<code>getFieldSchemaXmlById(fieldId: string, listTitle: string, context: IContext): Promise&lt;string&gt;</code>
</td>
<td>
Gets SchemaXml for the field by List Title and Field Id<br>
Parameters<br>
<code>fieldId</code> - Field's Id<br>
<code>listTitle</code> - List Title<br>
<code>context</code> - Customizer's context
</td>
</tr>
<tr>
<td>
<code>getPageViewId(context: IContext): string</code>
</td>
<td>
Gets correct view id from the page<br>
Parameters<br>
<code>context</code> - Customizer's context
</td>
</tr>
</tbody>
</table>

## Additional Information
The repository also contains Field Customizer to test the functionality
### Debug Url
?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&fieldCustomizers={"Percent":{"id":"57ebd944-98ed-43f9-b722-e959d6dac6ad"}}

## Release Notes
| Version | Description |
| ------- | ----------- |
| 1.0.4 | First stable release with all needed functionality |
| 1.1.0 | `window.g_listData` reference is completely removed<br>`SPHelper.getFieldSchemaXmlByTitleOrInternalName` is replaced with `SPHelper.getFieldSchemaXmlById`<br>`FieldRenderer.getFieldRenderer` is now asynchronous and shouldn't be called in `render` method of Field Customizer |
| 1.2.0 | All Field Controls names were prefixed with `Field`. As example, `FieldDateRenderer` instead of `DateRenderer`<br>Hyperlink or image rendering in Image mode is fixed<br>Lookup item display dialog is fixed |

## Contribution
Please, feel free to report any bugs or improvements for the repo.
If you're going to add a PR please, reference dev branch instead of master.

