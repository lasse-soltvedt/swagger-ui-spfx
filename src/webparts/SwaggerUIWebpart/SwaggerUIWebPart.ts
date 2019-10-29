import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as strings from 'SwaggerUiWebPartStrings';
import Wrapper from './components/Wrapper/Wrapper';
import IWrapperProps from './components/Wrapper/IWrapperProps';
import * as microsoftTeams from '@microsoft/teams-js';
import 'swagger-ui-react/swagger-ui.css';

export interface ISwaggerUIWebPartProps {
  url: string;
}

export default class SwaggerUIWebPart extends BaseClientSideWebPart<ISwaggerUIWebPartProps> {

  private _teamsContext: microsoftTeams.Context;
  private _isTeams: boolean = false;

  public render(): void {
    const element: React.ReactElement<IWrapperProps > = React.createElement(
      Wrapper,
      {
        url: this.properties.url,
        isTeams: this._isTeams
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onInit(): Promise<any> {

    // make spfx aware of teams context
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      this._isTeams = true;
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
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
                PropertyPaneTextField('url', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
