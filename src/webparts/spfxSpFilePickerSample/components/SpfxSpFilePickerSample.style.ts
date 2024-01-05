import { IProcessedStyleSet, mergeStyleSets } from '@fluentui/react/lib/Styling';

/**
 * Style dynamique
 */
export const classNames = (): IProcessedStyleSet<{ SPFilePicker: {}, FilePickerPanel: {} }> => {
    return mergeStyleSets({
        SPFilePicker: {
            width: '100%',
            height: '100%',
            border: 'none'
        },
        FilePickerPanel: {
            '.ms-Panel-content': {
                padding: 0
            },
            '.ms-Panel-commands': {
                display: 'none'
            },
            '.ms-Panel-scrollableContent': {
                overflow: 'hidden'
            }

        }
    });
};
