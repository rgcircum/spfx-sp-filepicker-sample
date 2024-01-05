// React
import * as React from 'react';
import { FC, useState } from 'react';

// Style
import { classNames } from './SpfxSpFilePickerSample.style';

// Fluent UI
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { useBoolean } from '@fluentui/react-hooks';
import { ActionButton } from '@fluentui/react/lib/Button';

// Interface / models / service
import { SPItem } from '../../../Models/IFilePicker';
import { ISpfxSpFilePickerSampleProps } from './ISpfxSpFilePickerSampleProps';
import FilePicker from './FilePickerSample';

const SpfxSpFilePickerSample: FC<ISpfxSpFilePickerSampleProps> = (props) => {

  const { serviceScope, theme, pickData, onPick } = props;

  const [_isPanelOpen, { setTrue: _openPanel, setFalse: _dismissPanel }] = useBoolean(false);
  const [_filePickerData, _setFilePickerData] = useState<SPItem[]>(pickData && pickData.length > 0 ? pickData : []);

  return (
    <>
      <div>
        <ActionButton
          text={'Open file picker'}
          iconProps={{ iconName: 'OpenFolderHorizontal' }}
          onClick={_openPanel}
        />
      </div>
      {_filePickerData && _filePickerData.length > 0 &&
        _filePickerData.map(data => {
          return (
            <div key={data.id}>
              {data?.webUrl}
            </div>
          )
        })
      }
      <Panel
        isOpen={_isPanelOpen}
        isBlocking={true}
        isLightDismiss={true}
        onDismiss={_dismissPanel}
        hasCloseButton={false}
        isFooterAtBottom={true}
        type={PanelType.large}
        className={classNames().FilePickerPanel}
      >
        <FilePicker
          pickData={pickData}
          serviceScope={serviceScope}
          theme={theme}
          onClose={_dismissPanel}
          onPick={(pickData: SPItem[]) => {
            _setFilePickerData(pickData);
            onPick(pickData);
            _dismissPanel()
          }}
        />
      </Panel>
    </>
  );
};

export default SpfxSpFilePickerSample;
