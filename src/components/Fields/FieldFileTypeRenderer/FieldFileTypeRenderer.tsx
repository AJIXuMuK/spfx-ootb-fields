import { override } from '@microsoft/decorators';
import * as React from 'react';
import { css, ISerializableObject, Icon } from 'office-ui-fabric-react';
import { IFieldRendererProps } from '../FieldCommon/IFieldRendererProps';
import { IconType, ApplicationIconList } from "@pnp/spfx-controls-react/lib/FileTypeIcon";

import styles from './FieldFileTypeRenderer.module.scss';
import { findIndex } from '@microsoft/sp-lodash-subset';

export interface IFieldFileTypeRendererProps extends IFieldRendererProps {
    /**
     * file/document path
     */
    path: string;
    /**
     * true if the icon should be rendered for a folder, not file
     */
    isFolder?: boolean;
}

/**
 * For future
 */
export interface IFieldFileTypeRendererState {

}

/**
 * File Type Renderer.
 * Used for:
 *   - File/Document Type
 */
export default class FieldFileTypeRenderer extends React.Component<IFieldFileTypeRendererProps, IFieldFileTypeRendererState> {
    public constructor(props: IFieldFileTypeRendererProps, state: IFieldFileTypeRendererState) {
        super(props, state);

        this.state = {};
    }

    @override
    public render(): JSX.Element {
        let iconName: string = '';
        if (this.props.isFolder) {
            iconName = 'FabricFolderFill';
        }
        else {
            const fileExtension: string = this._getFileExtension(this.props.path);
            iconName = this._getIconByExtension(fileExtension, IconType.font);
        }
        
        const optionalStyles: ISerializableObject = {   
        };
        optionalStyles[styles.folder] = this.props.isFolder;
        return (
            <div className={css(this.props.className, styles.container, styles.fabricIcon, optionalStyles)} style={this.props.cssProps}>
                <Icon iconName={iconName} />
            </div>
        );
    }

    /**
  * Function to retrieve the file extension from the path
  *
  * @param value File path
  */
  private _getFileExtension(value): string {
    // Split the URL on the dots
    const splittedValue = value.split('.');
    // Take the last value
    let extensionValue = splittedValue.pop();
    // Check if there are query string params in place
    if (extensionValue.indexOf('?') !== -1) {
      // Split the string on the question mark and return the first part
      const querySplit = extensionValue.split('?');
      extensionValue = querySplit[0];
    }
    return extensionValue;
  }

  /**
  * Find the icon name for the provided extension
  *
  * @param extension File extension
  */
  private _getIconByExtension(extension: string, iconType: IconType): string {
    // Find the application index by the provided extension
    const appIdx = findIndex(ApplicationIconList, item => { return item.extensions.indexOf(extension.toLowerCase()) !== -1; });

    // Check if an application has found
    if (appIdx !== -1) {
      // Check the type of icon, the image needs to get checked for the name
      if (iconType === IconType.font) {
        return ApplicationIconList[appIdx].iconName;
      } else {
        const knownImgs = ApplicationIconList[appIdx].imageName;
        // Check if the file extension is known
        const imgIdx = knownImgs.indexOf(extension);
        if (imgIdx !== -1) {
          return knownImgs[imgIdx];
        } else {
          // Return the first one if it was not known
          return knownImgs[0];
        }
      }
    }

    return 'Page';
  }
}