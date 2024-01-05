// React
import * as React from 'react';
import { FC, useCallback, useEffect, useMemo, useRef } from 'react';

// Style
import { classNames } from './SpfxSpFilePickerSample.style';

// SPFx
import { Guid } from '@microsoft/sp-core-library';

// Interface / models / service
import FilePickerService from '../../../Services/FilePicker.service';
import { EFoldersOrFiles, ESelectionMode, IFilePickerOptions, INotificationData, IPickData, PickerWindow } from '../../../Models/IFilePicker';
import { IFilePickerSampleProps } from './IFilePickerSampleProps';

const FilePicker: FC<IFilePickerSampleProps> = (props) => {

  const { serviceScope, theme, onPick, onClose } = props;

  const _iFrameRef: React.MutableRefObject<HTMLIFrameElement | null> = useRef(null);

  const _serviceInstance = useMemo(() => serviceScope.consume(FilePickerService.serviceKey), [serviceScope]);

  const _currentUICultureName = useMemo(() => _serviceInstance.getCurrentUICultureName(), [_serviceInstance]);
  const _webAbsoluteUrl = useMemo(() => _serviceInstance.getWebAbsoluteUrl(), [_serviceInstance]);

  const _channelId = useMemo(() => Guid.newGuid().toString(), []);

  // Config
  const _params: IFilePickerOptions = useMemo(() => ({
    sdk: '8.0',
    entry: {
      sharePoint: {
        byPath: {
          web: _webAbsoluteUrl,
          //folder: 'test',
          //list: 'internal name of list'
        }
      },
    },
    authentication: {},

    messaging: {
      origin: _webAbsoluteUrl,
      channelId: _channelId
    },
    selection: {
      mode: ESelectionMode.multiple,
    },
    search: {
      enabled: true
    },
    typesAndSources: {
      //filters: ['.docx'],
      mode: EFoldersOrFiles.all,
      pivots: {
        shared: true,
        oneDrive: true,
        recent: true,
        sharedLibraries: true
      },
    },
    accessibility: {
      focusTrap: 'initial'
    },

    commands: {
      close: {
        //label: 'close btn label',
      },
      pick: {
        action: 'share',
        //label: 'share btn label'
      },
    },
    localization: {
      language: _currentUICultureName
    },
    navigation: {},
    telemetry: {},
    theme: theme

  }), [_webAbsoluteUrl, _channelId, _currentUICultureName, theme]);

  const _onNotification = useCallback((e: CustomEvent<INotificationData>) => {
    console.log('picker notification: ', e.detail);
  }, []);

  const _onPickerChange = useCallback((e: CustomEvent<IPickData>) => {
    switch (e.detail.data.command) {
      case 'close':
        onClose();
        break;
      case 'pick':
        onPick(e.detail.data.items);
        break;
    }
  }, [onPick, onClose]);

  // LifeCycle
  useEffect(() => {

    let picker: PickerWindow;

    (async () => {
      if (_iFrameRef && _iFrameRef.current !== null && _iFrameRef.current.contentWindow) {

        picker = await _serviceInstance.picker(
          _iFrameRef.current.contentWindow,
          {
            type: 'ODSP',
            baseUrl: _webAbsoluteUrl,
            options: _params,
          }
        );

        if (DEBUG) picker.addEventListener('pickernotification', _onNotification);
        picker.addEventListener('pickerchange', _onPickerChange);
      }
    })().catch(console.error);

  }, [_onNotification, _onPickerChange, _params, _webAbsoluteUrl, _serviceInstance]);

  return (
    <iframe ref={_iFrameRef} title='browserFrame' id='browserFrame' className={classNames().SPFilePicker} />
  );
};

export default FilePicker;
