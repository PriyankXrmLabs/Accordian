import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneTextField
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
  list:string;
}

export default class AccordianWebPart extends BaseClientSideWebPart<IAccordianWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';


  public render(): void {
    const element: React.ReactElement<IAccordianProps> = React.createElement(
      Accordian,
      {
        description: this.properties.description,
        list: this.properties.list,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context:this.context,
        mode:this.displayMode
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
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
                PropertyPaneTextField('list', {
                  label: 'Enter List Name'
                }),
                PropertyPaneButton('button', {
                  text: 'Create List',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this._onButtonClick.bind(this)
                })
                

              ]
              
            }
          ]
        }
      ]
    };
  }


  private async _onButtonClick(): Promise<void> {
    try {
      const listName = this.properties.list; // Get the list name from the property pane
    
      if (!listName) {
        alert('Please enter a list name.');
        return;
      }
    
      const sp = spfi().using(SPFx(this.context));
    
      let listExists = false;
    
      try {
        // Check if the list exists
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
        // Create the list since it does not exist
        await sp.web.lists.add(listName, 'List created by SPFx WebPart', 100); // 100 is the default ListTemplateType
        const createdList = sp.web.lists.getByTitle(listName);
        await createdList.fields.addMultilineText('Description',{ RichText: false});
        alert(`List "${listName}" created successfully.`);
      } else {
        alert(`List "${listName}" already exists.`);
      }
    
    } catch (error) {
      console.log('Error creating list:', error);
      alert(error.message);
    }
    
  }
}
