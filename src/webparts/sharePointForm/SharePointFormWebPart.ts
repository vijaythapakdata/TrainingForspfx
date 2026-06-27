import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {sp} from "@pnp/sp/presets/all";

import * as strings from 'SharePointFormWebPartStrings';
import SharePointForm from './components/SharePointForm';
import { ISharePointFormProps } from './components/ISharePointFormProps';
import ChoiceFieldsClass from '../../api/services/ChoiceServiceApi';

export interface ISharePointFormWebPartProps {
  description: string;
}

export default class SharePointFormWebPart extends BaseClientSideWebPart<ISharePointFormWebPartProps> {

 private choicesOptions:ChoiceFieldsClass|undefined;

  protected async onInit(): Promise<void> {
    this.choicesOptions=new ChoiceFieldsClass(this.context);
    return super.onInit().then(_=> {
     sp.setup({
      spfxContext:this.context as any
     });
    });
  }

  public async render(): Promise<void> {
    const element: React.ReactElement<ISharePointFormProps> = React.createElement(
      SharePointForm,
      {
        context:this.context,
        siteurl:this.context.pageContext.web.absoluteUrl,
        singleselecteddropdown:await this.choicesOptions?.getChoiceValues(this.context.pageContext.web.absoluteUrl,"Department"),
      radiooption:await this.choicesOptions?.getChoiceValues(this.context.pageContext.web.absoluteUrl,"Gender"),
      multiselectdropdown:await this.choicesOptions?.getChoiceValues(this.context.pageContext.web.absoluteUrl,"Skills"),
      lookupval:await this.choicesOptions?.getLookupValues()
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
