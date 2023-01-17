import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { PropertyPaneWebPartInformation } from "@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation";
import * as strings from "SimpleBannerWebPartStrings";
import SimpleBanner from "./components/SimpleBanner";
import { ISimpleBannerProps } from "./components/ISimpleBannerProps";

export interface ISimpleBannerWebPartProps {
  description: string;
  itemId: number;
  fileName: string;
  fileSize: number;
}

export default class SimpleBannerWebPart extends BaseClientSideWebPart<ISimpleBannerWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ISimpleBannerProps> = React.createElement(
      SimpleBanner,
      {
        context: this.context,
        itemId: this.properties.itemId,
        updatePropety: (id: number) => {
          this.properties.itemId = id;
        },
        fileName: this.properties.fileName,
        fileSize: this.properties.fileSize,
        updateFileName: (filename: string) => {
          this.properties.fileName = filename;
        },
        updateFileSize: (filesize: number) => {
          this.properties.fileSize = filesize;
        },
      }
    );
    ReactDom.render(element, this.domElement);
  }

  onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.descriptionImage,
              isCollapsed: true,
              groupFields: [
                PropertyPaneWebPartInformation({
                  description: this.properties.fileName
                    ? `<div style="font-size: 1rem;">${strings.fileName}:</div><div style="font-size: 1rem;"><strong>${this.properties.fileName}</strong></div>
                    <div style="font-size: 1rem;">${strings.fileSize}:</div><div style="font-size: 1rem;"><strong>${this.properties.fileSize}</strong></div>
                    <div style="font-size: 1rem;">ID:</div><div><strong>${this.properties.itemId}</strong></div>`
                    : `${strings.addImage}`,
                  key: "webPartInfoId",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
