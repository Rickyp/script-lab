import * as $ from 'jquery';
import { isNil } from 'lodash';
import { UI } from '@microsoft/office-js-helpers';
import { environment, instantiateRibbon, generateUrl, navigateToRegisterCustomFunctions, displayLogDialog } from '../app/helpers';
import { Strings, setDisplayLanguage } from '../app/strings';
import { officeNamespacesForIframe } from './runner.common';
import { Messenger, CustomFunctionsMessageType } from '../app/helpers/messenger';

interface InitializationParams {
    snippetDataBase64: string;
    explicitlySetDisplayLanguageOrNull: string;
    returnUrl: string;
}

const CSS_CLASSES = {
    inProgress: 'in-progress',
    error: 'error',
    success: 'success'
};

let allSuccessful = true;

// Note: Office.initialize is already handled outside in the html page,
// setting "window.playground_host_ready = true;"

(async () => {
    let params: InitializationParams = (window as any).customFunctionParams;

    try {
        environment.initializePartial({ host: 'EXCEL' });

        // Apply the host theming by adding this attribute on the "body" element:
        $('body').addClass('EXCEL');
        $('#header').css('visibility', 'visible');

        if (instantiateRibbon('ribbon')) {
            $('#progress').css('border-top', '#ddd 5px solid;');
        }

        await new Promise((resolve) => {
            const interval = setInterval(() => {
                if ((window as any).playground_host_ready) {
                    clearInterval(interval);
                    return resolve();
                }
            }, 100);
        });

        // Set initialize to an empty function -- that way, doesn't cause
        // re-initialization of this page in case of a page like the error dialog,
        // which doesn't defined (override) Office.initialize.
        Office.initialize = () => { };

        await initializeRegistration(params);
    }
    catch (error) {
        handleError(error);
    }
})();

async function initializeRegistration(params: InitializationParams) {
    if (params.explicitlySetDisplayLanguageOrNull) {
        setDisplayLanguage(params.explicitlySetDisplayLanguageOrNull);
        document.cookie = `displayLanguage=${encodeURIComponent(params.explicitlySetDisplayLanguageOrNull)};path=/;`;
    }

    const $snippetNames = $('#snippet-names');

    const snippetData: ICustomFunctionsRegistrationRelevantData[] = JSON.parse(atob(params.snippetDataBase64));
    const actualCount = snippetData.length - 1;
    /* Last one is always null, set in the template for ease of trailing commas... */

    for (let i = 0; i < actualCount; i++) {
    // MAJOR FIXME
    //     const snippetBase64OrNull = params.snippetIframesBase64Texts[i];
    //     let $entry = showUI ? $snippetNames.children().eq(i) : null;

    //     if (isNil(snippetBase64OrNull) || snippetBase64OrNull.length === 0) {
    //         $entry.addClass(CSS_CLASSES.error);
            
    //         allSuccessful = false;
    //     }
    //     else {
    //         if (showUI) {
    //             $entry.addClass(CSS_CLASSES.inProgress);
    //         }

    //         let success = await runSnippetCode(atob(params.snippetIframesBase64Texts[i]));
    //         if (showUI) {
    //             $entry.removeClass(CSS_CLASSES.inProgress)
    //                 .addClass(success ? CSS_CLASSES.success : CSS_CLASSES.error);
    //         } else {
    //             if (success) {
    //                 heartbeat.messenger.send<LogData>(heartbeat.window, CustomFunctionsMessageType.LOG, {
    //                     timestamp: new Date().getTime(),
    //                     source: 'system',
    //                     type: 'custom functions',
    //                     subtype: 'runner',
    //                     severity: 'info',
    //                     // TODO CUSTOM FUNCTIONS localization
    //                     message: `Sucessfully loaded "${params.snippetNames[i]}"`
    //                 });
    //             } else {
    //                 heartbeat.messenger.send<LogData>(heartbeat.window, CustomFunctionsMessageType.LOG, {
    //                     timestamp: new Date().getTime(),
    //                     source: 'system',
    //                     type: 'custom functions',
    //                     subtype: 'runner',
    //                     severity: 'error',
    //                     // TODO CUSTOM FUNCTIONS localization
    //                     message: `Could NOT load function "${params.snippetNames[i]}"`
    //                 });
    //             }
    //         }

    //         allSuccessful = allSuccessful && success;

    //     }
    // }

    // // Complete any function registrations
    // await Excel.run(async (context) => {
    //     (context.workbook as any).customFunctions.addAll();
    //     await context.sync();
    // });

    // if (showUI && !allSuccessful) {
    //     $('.ms-progress-component__footer').css('visibility', 'hidden');
    // }

    // if (isRunMode) {
    //     heartbeat.messenger.send<{ timestamp: number }>(heartbeat.window,
    //         CustomFunctionsMessageType.LOADED_AND_RUNNING, { timestamp: new Date().getTime() });
    // }
    // else {
    //     if (allSuccessful) {
    //         window.location.href = params.returnUrl;
    //     }
    // }
}

function handleError(error: Error) {
    allSuccessful = false;

    let candidateErrorString = error.message || error.toString();
    if (candidateErrorString === '[object Object]') {
        candidateErrorString = Strings().unexpectedError;
    }

    if (error instanceof Error) {
        UI.notify(error);
    } else {
        UI.notify(Strings().error, candidateErrorString);
    }
}

