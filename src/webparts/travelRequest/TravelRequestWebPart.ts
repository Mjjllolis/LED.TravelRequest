import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';

import * as strings from 'TravelRequestWebPartStrings';
import TravelRequest from './components/TravelRequest';
import { ITravelRequestProps } from './components/ITravelRequestProps';
import { sp } from '@pnp/sp';
import './components/TravelRequest.module.print.css';

export interface ITravelRequestWebPartProps {
  mileageRate: number;
  startingFinancialYear: number;
}

export default class TravelRequestWebPart extends BaseClientSideWebPart<ITravelRequestWebPartProps> {
  public onInit(): Promise<void> {
    //configure pnp/sp
    return super.onInit().then(_ => {
      // other init code may be present
      sp.setup({
        spfxContext: this.context
      });     
    });
  }
  public render(): void {
    const element: React.ReactElement<ITravelRequestProps> = React.createElement(
      TravelRequest,
      {
        mileageRate: this.properties.mileageRate,
        context:this.context,
        startingFinancialYear: this.properties.startingFinancialYear,
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
                PropertyPaneTextField('mileageRate', {
                  label: 'Default Mileage Rate'
                }),
                PropertyPaneTextField('startingFinancialYear', {
                  label: 'Default Starting Financial Year'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
