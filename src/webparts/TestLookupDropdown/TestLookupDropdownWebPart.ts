/*
  tslint:disable:max-line-length no-any
*/
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration
} from '@microsoft/sp-webpart-base';

import TestLookupDropdown from './components/TestLookupDropdown/TestLookupDropdown';
import { ITestLookupDropdownProps } from './components/TestLookupDropdown/ITestLookupDropdownProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls';
import { sp } from '@pnp/sp';

export interface ITestLookupDropdownWebPartProps {
  list1: string;
  list2: string;
  list3: string;
}

export default class TestLookupDropdownWebPart extends BaseClientSideWebPart<ITestLookupDropdownWebPartProps> {

  constructor() {
    super();
    sp.setup({
      spfxContext: this.context
    });
  }

  public render(): void {
    const element: React.ReactElement<ITestLookupDropdownProps> = React.createElement(
      TestLookupDropdown,
      {
        listLand: this.properties.list1,
        listStadt: this.properties.list2,
        listStrasse: this.properties.list3
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /// @ts-ignore
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected checkStadtListIsValid = async (listId: string, parentListId: string): Promise<string> => {
    if (listId !== undefined && listId !== '') {
      try {
        const field: any = await sp.web.lists.getById(listId).fields.getByInternalNameOrTitle('Land').get();
        if (field.LookupList === `{${parentListId}}`) {
          return '';
        } else {
          return 'Die Spalte \'Land\' verweist nicht auf die erste Liste!';
        }
      } catch (error) {
        return 'Die Liste enthält keine Spalte mit Namen \'Land\'.';
      }
    } else {
      return '';
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Einstellungen'
          },
          groups: [
            {
              groupName: 'Listen auswählen',
              groupFields: [
                PropertyFieldListPicker('list1', {
                  label: 'Erste Liste',
                  selectedList: this.properties.list1,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: undefined,
                  deferredValidationTime: 0,
                  key: 'list1Id'
                }),
                PropertyFieldListPicker('list2', {
                  label: 'Zweite Liste',
                  selectedList: this.properties.list2,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: (list => this.checkStadtListIsValid(list, this.properties.list1)),
                  deferredValidationTime: 0,
                  key: 'list2Id'
                }),
                PropertyFieldListPicker('list3', {
                  label: 'Dritte Liste',
                  selectedList: this.properties.list3,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: undefined,
                  deferredValidationTime: 0,
                  key: 'list3Id'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
