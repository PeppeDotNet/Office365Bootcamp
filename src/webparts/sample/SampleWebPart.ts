import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as strings from 'SampleWebPartStrings';
import Sample from './components/Sample';
import { ISampleProps } from './components/Sample';
import { PropertyFieldSPListPicker, PropertyFieldSPListPickerOrderBy } from 'sp-client-custom-fields/lib/PropertyFieldSPListPicker';

export interface ISampleWebPartProps {
  list: string;
}

export default class SampleWebPart extends BaseClientSideWebPart<ISampleWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISampleProps > = React.createElement(
      Sample,
      {
        list: this.properties.list,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyFieldSPListPicker('list', {
                  label: 'Select a contact list',
                  selectedList: this.properties.list,
                  includeHidden: false,
                  baseTemplate: 105,
                  orderBy: PropertyFieldSPListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this.disableReactivePropertyChanges,
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 200,
                  key: 'listPickerFieldId'
               })
              ]
            }
          ]
        }
      ]
    };
  }
}
