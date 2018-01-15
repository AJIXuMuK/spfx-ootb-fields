import * as React from "react";
import * as ReactDOM from "react-dom";

import { BaseDialog, IDialogConfiguration } from "@microsoft/sp-dialog";
import IFrameDialogContent from "./IFrameDialogContent";

/**
 * Dialog component to display content in iframe
 */
export default class IFrameDialog extends BaseDialog {
    public url: string;
    public iframeOnLoad: (iframe: any) => {};
  
    public render(): void {
      ReactDOM.render(<IFrameDialogContent url={this.url} close={this.close} iframeOnLoad={this.iframeOnLoad} />, this.domElement);
    }
  
    public getConfig(): IDialogConfiguration {
      return {
        isBlocking: true
      };
    }
  }