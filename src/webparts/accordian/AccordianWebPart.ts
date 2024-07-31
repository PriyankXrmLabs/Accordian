import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneChoiceGroup,
  IPropertyPaneDropdownOption,
  IPropertyPaneField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'AccordianWebPartStrings';
import Accordian from './components/Accordian';
import { IAccordianProps } from './components/IAccordianProps';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";

export interface IAccordianWebPartProps {
  description: string;
  listChoice: string;
  listName: string;
  selectionOption: string;
}

export default class AccordianWebPart extends BaseClientSideWebPart<IAccordianWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _lists: IPropertyPaneDropdownOption[] = [];

  public async onInit(): Promise<void> {
    await this._fetchLists();
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private async _fetchLists(): Promise<void> {
    const sp = spfi().using(SPFx(this.context));
    try {
      const lists = await sp.web.lists.select('Title')();
      this._lists = lists.map(list => ({ key: list.Title, text: list.Title }));
      this.context.propertyPane.refresh();
    } catch (error) {
      console.error('Error fetching lists:', error);
    }
  }

  public render(): void {
    const element: React.ReactElement<IAccordianProps> = React.createElement(
      Accordian,
      {
        description: this.properties.description,
        list: this.properties.listName,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        mode: this.displayMode
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
  
    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const fields: IPropertyPaneField<any>[] = [
      PropertyPaneTextField('description', {
        label: strings.DescriptionFieldLabel
      }),
      PropertyPaneChoiceGroup('selectionOption', {
        label: 'Select an option',
        options: [
          { key: 'create', text: 'Create List' },
          { key: 'select', text: 'Select Existing List' }
        ],
        
      })
    ];

    if (this.properties.selectionOption === 'create') {
      fields.push(PropertyPaneTextField('listName', {
        label: 'Enter New List Name'
      }));
    } else if (this.properties.selectionOption === 'select') {
      fields.push(PropertyPaneDropdown('listChoice', {
        label: 'Select Existing List',
        options: this._lists,
        selectedKey: this.properties.listChoice
      }));
    }

    fields.push(PropertyPaneButton('button', {
      text: 'Apply',
      buttonType: PropertyPaneButtonType.Primary,
      onClick: this._onButtonClick.bind(this)
    }));

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: fields
            }
          ]
        }
      ]
    };
  }

  private async _onButtonClick(): Promise<void> {
    try {
      const { selectionOption, listChoice, listName } = this.properties;

      if (selectionOption === 'create') {
        if (!listName) {
          alert('Please enter a list name.');
          return;
        }
      
        await this._createNewList(listName);
      } else if (selectionOption === 'select') {
        alert(`You selected the pre-built list: ${listChoice}`);
        this.properties.listName = listChoice;
      }
    } catch (error) {
      console.log('Error:', error);
      alert(error.message);
    }
  }

  private async _createNewList(listName: string): Promise<void> {
    const sp = spfi().using(SPFx(this.context));

    let listExists = false;

    try {
      await sp.web.lists.getByTitle(listName).select('Title')();
      listExists = true;
    } catch (error) {
      if (error.message.includes("404")) {
        listExists = false;
      } else {
        throw error;
      }
    }

    if (!listExists) {
      await sp.web.lists.add(listName, 'List created by SPFx WebPart', 100);
      const createdList = sp.web.lists.getByTitle(listName);
      await createdList.fields.addMultilineText('Description', { RichText: false });
      alert(`List "${listName}" created successfully.`);
    } else {
      alert(`List "${listName}" already exists.`);
    }
  }
}
