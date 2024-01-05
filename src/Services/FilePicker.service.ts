/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-case-declarations */
import { AadTokenProviderFactory } from '@microsoft/sp-http';
import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';

import {
    IAuthenticateCommand,
    INotificationData,
    IPickData,
    ODSPInit,
    OneDriveConsumerInit,
    PickerWindow,
} from '../Models/IFilePicker';

import { combine } from '../Helpers/utils';

export interface IFilePickerService {
    picker(win: Window, init: OneDriveConsumerInit | ODSPInit): Promise<PickerWindow>;
    getToken(command: IAuthenticateCommand): Promise<string>;
    getCurrentUICultureName(): string;
    getWebAbsoluteUrl(): string;
}

const EXP_SOURCE: string = 'SpfxSpFilePickerSampleWebPart::PickerService';

export default class FilePickerService implements IFilePickerService {

    public static readonly serviceKey: ServiceKey<IFilePickerService> = ServiceKey.create<IFilePickerService>(EXP_SOURCE, FilePickerService);

    private _pageContext: PageContext;
    private _aadTokenProviderFactory: AadTokenProviderFactory;

    /**
     * Service constructor
     */
    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
            this._aadTokenProviderFactory = serviceScope.consume(AadTokenProviderFactory.serviceKey);
        });
    }

    /**
     * Get current UI culture name
     * @returns 
     */
    public getCurrentUICultureName(): string {
        return this._pageContext.cultureInfo.currentUICultureName;
    }

    /**
    * Get absolute url
    * @returns 
    */
    public getWebAbsoluteUrl(): string {
        return this._pageContext.web.absoluteUrl;
    }

    /**
     * Get token
     * @param command 
     * @returns 
     */
    public async getToken(command: IAuthenticateCommand): Promise<string> {
        const _aadTokenProvider = await this._aadTokenProviderFactory.getTokenProvider();
        const _authToken = await _aadTokenProvider.getToken(command.resource);
        return _authToken;
    }

    /**
     * Initializes and loads a new file picker into the provided window using the supplied init
     * @param win The window (iframe) into which the file picker will be loaded
     * @param init The initialization used to create the file picker
     * @returns A picker window interface
     */
    public async picker(win: Window, init: OneDriveConsumerInit | ODSPInit): Promise<PickerWindow> {

        // this is the port we'll use to communicate with the picker
        let port: MessagePort;

        // we default to the consumer values since they are fixed
        const baseUrl = init.type === 'ODSP' ? init.baseUrl : 'https://onedrive.live.com';
        const pickerPath = combine(baseUrl, init.type === 'ODSP' ? '_layouts/15/FilePicker.aspx' : 'picker');

        // grab the things we need from the init
        const { options } = init;

        // eslint-disable-next-line @typescript-eslint/no-this-alias
        const that = this;

        // define the message listener to process the various messages from the window
        async function messageListener(message: MessageEvent): Promise<void> {

            switch (message.data.type) {

                case 'notification':

                    const notification = message.data;

                    if (notification.notification === 'page-loaded') {
                        console.log('page-loaded');
                        // here we know that the picker page is loaded and ready for user interaction
                    }

                    window.dispatchEvent(new CustomEvent<INotificationData>('pickernotification', {
                        detail: message.data
                    }));

                    break;

                case 'command':

                    port.postMessage({
                        type: 'acknowledge',
                        id: message.data.id,
                    });

                    const command = message.data.data;

                    switch (command.command) {

                        case 'authenticate':

                            const token = await that.getToken(command);

                            if (typeof token !== 'undefined') {

                                port.postMessage({
                                    type: 'result',
                                    id: message.data.id,
                                    data: {
                                        result: 'token',
                                        token,
                                    },
                                });
                            }

                            break;

                        case 'close':

                            window.dispatchEvent(new CustomEvent<IPickData>('pickerchange', {
                                detail: message.data
                            }));

                            port.postMessage({
                                type: 'result',
                                id: message.data.id,
                                data: {
                                    result: 'success',
                                },
                            });

                            break;

                        case 'pick':

                            window.dispatchEvent(new CustomEvent<IPickData>('pickerchange', {
                                detail: message.data
                            }));

                            port.postMessage({
                                type: 'result',
                                id: message.data.id,
                                data: {
                                    result: 'success',
                                },
                            });

                            break;

                        default:

                            console.warn(`Unsupported picker command: ${JSON.stringify(command)}`);

                            // let the picker know we don't support whatever command it sent
                            port.postMessage({
                                result: 'error',
                                error: {
                                    code: 'unsupportedCommand',
                                    message: command.command
                                },
                                isExpected: true,
                            });
                            break;
                    }

                    break;
            }
        }

        // attach a listener for the message event to setup our channel
        window.addEventListener('message', (event) => {
            if (event.source && event.source === win) {
                const message = event.data;
                if (message.type === 'initialize' && message.channelId === options.messaging?.channelId) {
                    port = event.ports[0];
                    port.addEventListener('message', messageListener);
                    port.start();
                    port.postMessage({
                        type: 'activate',
                    });
                }
            }
        });

        const authToken = await this.getToken({
            command: 'authenticate',
            type: 'SharePoint',
            resource: baseUrl,
        });

        const queryString = new URLSearchParams({
            filePicker: JSON.stringify(options),
        });

        const url = `${pickerPath}?${queryString}`;

        // now we post a form into the window to load the picker with the options
        const form = win.document.createElement('form');
        form.setAttribute('action', url);
        form.setAttribute('method', 'POST');
        win.document.body.append(form);

        if (authToken !== null) {

            const input = win.document.createElement('input');
            input.setAttribute('type', 'hidden');
            input.setAttribute('name', 'access_token');
            input.setAttribute('value', authToken);
            form.appendChild(input);
        }

        // this will load the picker into the window
        form.submit();

        // we return the current global window, which will get sent the custom events
        // when there are notifications or items are picked, but we scoped down the typings
        // to make intendend options clearer
        return window;
    }
}