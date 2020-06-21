import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PnpPropertyFieldCollectionDataWebPartStrings';
import PnpPropertyFieldCollectionData from './components/PnpPropertyFieldCollectionData';
import { IPnpPropertyFieldCollectionDataProps } from './components/IPnpPropertyFieldCollectionDataProps';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

export interface IPnpPropertyFieldCollectionDataWebPartProps {
  description: string;
  collectionData: any[];
}

export default class PnpPropertyFieldCollectionDataWebPart extends BaseClientSideWebPart<IPnpPropertyFieldCollectionDataWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPnpPropertyFieldCollectionDataProps> = React.createElement(
      PnpPropertyFieldCollectionData,
      {
        description: this.properties.description,
        collectionData: this.properties.collectionData
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldCollectionData("collectionData", {
                  key: "collectionData",
                  label: "Collection data",
                  panelHeader: "Collection data panel header",
                  manageBtnLabel: "Manage collection data",
                  value: this.properties.collectionData,
                  fields: [
                    {
                      id: "Title",
                      title: "Firstname",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "Lastname",
                      title: "Lastname",
                      type: CustomCollectionFieldType.string
                    },
                    {
                      id: "Age",
                      title: "Age",
                      type: CustomCollectionFieldType.number,
                      required: true
                    },
                    {
                      id: "City",
                      title: "Favorite city",
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        {
                          key: "pune",
                          text: "Pune"
                        },
                        {
                          key: "junagadh",
                          text: "Junagadh"
                        }
                      ],
                      required: true
                    },
                    {
                      id: "Sign",
                      title: "Signed",
                      type: CustomCollectionFieldType.boolean
                    }, {
                      id: "customFieldId",
                      title: "People picker",
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                        return (
                          React.createElement(PeoplePicker, {
                            context: this.context,
                            personSelectionLimit: 1,
                            showtooltip: true,
                            key:itemId,
                            defaultSelectedUsers:[item.customFieldId],
                            selectedItems: (items: any[]) => {
                              console.log('Items:', items);
                              item.customFieldId = items[0].secondaryText;
                              onUpdate(field.id, items[0].secondaryText);
                            },
                            showHiddenInUI: false,
                            principalTypes: [PrincipalType.User]
                          }
                          )
                        );
                      }
                    }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
