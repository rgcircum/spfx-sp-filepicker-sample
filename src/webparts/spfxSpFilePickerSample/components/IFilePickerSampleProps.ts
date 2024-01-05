import { ServiceScope } from '@microsoft/sp-core-library';
import { SPItem } from '../../../Models/IFilePicker';
import { ITheme } from '@fluentui/react/lib/Styling';

export interface IFilePickerSampleProps {
  serviceScope: ServiceScope;
  theme?: ITheme;
  pickData : SPItem[]
  onPick : (pickData : SPItem[])=> void;
  onClose : ()=> void;
}
