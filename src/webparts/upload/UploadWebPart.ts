import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'UploadWebPartStrings';
import Upload from './components/Upload';
import { IUploadProps } from './components/IUploadProps';

export interface IUploadWebPartProps {
  ListName: string;
}

export default class UploadWebPart extends BaseClientSideWebPart<IUploadWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IUploadProps> = React.createElement(
      Upload,
      {
        ListName: this.properties.ListName,
        context:this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('ListName', {
                  label: 'Enter List Name'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
