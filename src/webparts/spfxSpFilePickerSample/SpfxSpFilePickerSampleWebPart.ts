import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';
import SpfxSpFilePickerSample from './components/SpfxSpFilePickerSample';
import { ISpfxSpFilePickerSampleProps } from './components/ISpfxSpFilePickerSampleProps';
import { SPItem } from '../../Models/IFilePicker';
import { ITheme } from '@fluentui/react/lib/Styling';

export interface ISpfxSpFilePickerSampleWebPartProps {
  pickData: SPItem[]
}

export default class SpfxSpFilePickerSampleWebPart extends BaseClientSideWebPart<ISpfxSpFilePickerSampleWebPartProps> {

  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  protected async onInit(): Promise<void> {

    await super.onInit();

    this.context.serviceScope.whenFinished(() => {

      // Theme
      this._themeProvider = this.context.serviceScope.consume(
        ThemeProvider.serviceKey
      );
      this._themeVariant = this._themeProvider.tryGetTheme();
      this._themeProvider.themeChangedEvent.add(
        this,
        this._handleThemeChangedEvent
      );
    });

    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<ISpfxSpFilePickerSampleProps> = React.createElement(
      SpfxSpFilePickerSample,
      {
        theme: this._themeVariant as ITheme,
        serviceScope: this.context.serviceScope,
        pickData: this.properties.pickData,
        onPick: (pickData: SPItem[]) => {
          this.properties.pickData = pickData
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme as ITheme;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
