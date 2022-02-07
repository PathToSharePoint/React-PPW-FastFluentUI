import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneHorizontalRule,
  PropertyPaneLabel,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';
import { update } from '@microsoft/sp-lodash-subset';
import { PropertyPaneWrap } from 'property-pane-wrap';

import { provideFluentDesignSystem, fluentTextField, fluentRadio, fluentRadioGroup, fluentSlider, fluentSliderLabel, fluentSwitch, StandardLuminance, fluentSelect, fluentOption, fluentDesignSystemProvider, baseLayerLuminance } from '@fluentui/web-components';
import { provideReactWrapper } from '@microsoft/fast-react-wrapper';
import { css } from "@microsoft/fast-element";

const fluentReactWrapper = provideReactWrapper(React);

const FluentTextField = fluentReactWrapper.wrap(fluentTextField());
const FluentRadio = fluentReactWrapper.wrap(fluentRadio());
const FluentRadioGroup = fluentReactWrapper.wrap(fluentRadioGroup());
const FluentSlider = fluentReactWrapper.wrap(fluentSlider());
const FluentSliderLabel = fluentReactWrapper.wrap(fluentSliderLabel());
const FluentSwitch = fluentReactWrapper.wrap(fluentSwitch());
const FluentSelect = fluentReactWrapper.wrap(fluentSelect());
const FluentOption = fluentReactWrapper.wrap(fluentOption());

provideFluentDesignSystem().register(fluentReactWrapper.registry);

export interface IHelloWorldWebPartProps {
  description: string;
  fluentMode: any;
  fluentTextField: string;
  fluentSlider: string;
  fluentSwitch: string;
  fluentRadioGroup: string;
  fluentSelect: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        properties: this.properties,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public updateWebPartProperty(property, value, refreshWebPart = true, refreshPropertyPane = true) {

    update(this.properties, property, () => value);
    if (refreshWebPart) this.render();
    if (refreshPropertyPane) this.context.propertyPane.refresh();

  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    let fluentMode;
    switch (this.properties.fluentMode) {
      case "inherit": fluentMode = this._isDarkTheme ? StandardLuminance.DarkMode : StandardLuminance.LightMode; break;
      default: fluentMode = this.properties.fluentMode;
    }

    if (fluentMode) baseLayerLuminance.setValueFor(document.body, fluentMode);

    const familiesList = [
      { name: "Office", value: 'Office' },
      { name: "M365", value: 'M365' }
    ];
    const appsList = [
      { header: "Teams", key: 'Teams', parent: "M365" },
      { header: "OneDrive", key: 'OneDrive', parent: "M365" },
      { header: "Yammer", key: 'Yammer', parent: "M365" },
      { header: "Excel", key: 'Excel', parent: "Office" },
      { header: "PowerPoint", key: 'PowerPoint', parent: "Office" },
      { header: "Word", key: 'Word', parent: "Office" }
    ];

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Fast Fluent UI Controls",
              groupFields: [
                PropertyPaneHorizontalRule(),
                PropertyPaneChoiceGroup("fluentMode", {
                  options: [
                    { key: "inherit", text: "Infer from context" },
                    { key: StandardLuminance.DarkMode, text: "Dark Mode" },
                    { key: StandardLuminance.LightMode, text: "Light Mode" }
                  ],
                  label: "Fast Fluent UI Mode"
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLabel("fluentSlider", { text: "Fast Fluent UI Slider" }),
                PropertyPaneWrap("fluentSlider", {
                  component: () =>
                    <FluentSlider
                      min="0"
                      max="100"
                      step="5"
                      title="Set the temperature"
                      value={this.properties.fluentSlider}
                      onClick={(e) => this.updateWebPartProperty("fluentSlider", e.target.value)}
                    >
                      {/* <FluentSliderLabel position="0">0 &#8451;</FluentSliderLabel>
                      <FluentSliderLabel position="25" >25 &#8451;</FluentSliderLabel>
                      <FluentSliderLabel position="50" >50 &#8451;</FluentSliderLabel>
                      <FluentSliderLabel position="75" >75 &#8451;</FluentSliderLabel>
                      <FluentSliderLabel position="100" >100 &#8451;</FluentSliderLabel> */}
                    </FluentSlider>
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLabel("fluentRadioGroup", { text: "Fast Fluent UI Radio Group" }),
                PropertyPaneWrap("fluentRadioGroup", {
                  component: () =>
                    <FluentRadioGroup
                      value={this.properties.fluentRadioGroup}
                      orientation="horizontal"
                      onClick={(e) => {
                        this.updateWebPartProperty("fluentRadioGroup", e.target.value);
                        this.updateWebPartProperty("fluentSelect", " ");
                      }}
                    >
                      {familiesList.map(item => <FluentRadio value={item.value}>{item.name}</FluentRadio>)}
                    </FluentRadioGroup>
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLabel("fluentSelect", { text: "Fast Fluent UI Select (Cascading)" }),
                PropertyPaneWrap("fluentSelect", {
                  component: () =>
                    <FluentRadioGroup
                      value={this.properties.fluentSelect}
                      orientation="vertical"
                      onClick={(e) => {
                        this.updateWebPartProperty("fluentSelect", e.target.value);
                      }}
                    >
                      {appsList
                        .filter(i => i.parent == this.properties["fluentRadioGroup"])
                        .map(item => <FluentRadio value={item.key}>{item.header}</FluentRadio>)}
                    </FluentRadioGroup>
                  // <FluentSelect
                  // defaultValue={this.properties.fluentSelect}
                  // style={{ width: "100%" }}
                  //   onInput={(e) => this.updateWebPartProperty("fluentSelect", e.target.currentValue)}
                  // >
                  //   <FluentOption value=" ">Select an app...</FluentOption>
                  //   {appsList
                  //     .filter(i => i.parent == this.properties["fluentRadioGroup"])
                  //     .map(item => <FluentOption selected={(item.key==this.properties.fluentSelect)} value={item.key}>{item.header}</FluentOption>)
                  //   }
                  // </FluentSelect>
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLabel("fluentTextField", { text: "Fast Fluent UI Text Field" }),
                PropertyPaneWrap("fluentTextField", {
                  component: FluentTextField,
                  props: {
                    value: this.properties.fluentTextField,
                    style: { width: "100%" },
                    onInput: (e) => this.updateWebPartProperty("fluentTextField", e.target.value)
                  }
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneLabel("fluentSwitch", { text: "Fast Fluent UI Switch" }),
                PropertyPaneWrap("fluentSwitch", {
                  component: FluentSwitch,
                  props: {
                    checked: this.properties.fluentSwitch,
                    onClick: (e) => this.updateWebPartProperty("fluentSwitch", e.target.checked)
                  }
                })
              ]
            },
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneHorizontalRule(),
                PropertyPaneTextField('description', {
                  label: strings.PropertyPaneTextFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
