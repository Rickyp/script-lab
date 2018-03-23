import * as $ from 'jquery';
import { UI } from '@microsoft/office-js-helpers';
import { environment, instantiateRibbon } from '../app/helpers';
import { Strings, setDisplayLanguage } from '../app/strings';

interface InitializationParams {
    snippetsDataBase64: string;
    explicitlySetDisplayLanguageOrNull: string;
    returnUrl: string;
}


// Note: Office.initialize is already handled outside in the html page,
// setting "window.playground_host_ready = true;""

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

    const snippetsDataArray: ICustomFunctionsRegistrationRelevantData[] = JSON.parse(atob(params.snippetsDataBase64));

    let customFunctionsMetadata: ICustomFunctionsRegistrationApiMetadata = validatesnippetsDataArray(snippetsDataArray);

    // Complete any function registrations
    await Excel.run(async (context) => {
        context.workbook.worksheets.getActiveWorksheet().getCell(0, 0).values = [[JSON.stringify(customFunctionsMetadata, null, 4)]];
        await context.sync();
    });

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

function validatesnippetsDataArray(snippetsDataArray: ICustomFunctionsRegistrationRelevantData[]): ICustomFunctionsRegistrationApiMetadata {
    let customFunctionsMetadata: ICustomFunctionsRegistrationApiMetadata = {
        functions: []
    };

    // In the excel side the parser only cared for the required and optional parameters, any extra stuff is jsut ignored, as long as each function contains the required parameters it is okay any info they send
    // None the less, the validation of duplicate names does sound like a logic thing to be done here.

    const snippetDictionary: { [key: string]: boolean } = {};

    snippetsDataArray.forEach(currentSnippet => {
        const functionDictionary: { [key: string]: boolean } = {};

        if (snippetDictionary[currentSnippet.name]) {
            // Duplicate found, deal with it and break
            // TODO ricky
        }
        snippetDictionary[currentSnippet.name] = true;

        currentSnippet.data.functions.forEach(currentFunction => {
            const parameterDictionary: { [key: string]: boolean } = {};

            if (functionDictionary[currentFunction.name]) {
                // Duplicate found, deal with it and break
                // TODO
            }
            functionDictionary[currentFunction.name] = true;

            currentFunction.parameters.forEach(currentParameter => {
                if (parameterDictionary[currentParameter.name]) {
                    // Duplicate found, deal with it and break
                    // TODO
                }
                parameterDictionary[currentParameter.name] = true;
            });

            //The function is OKAY, append namespace to function name so that it can be the real custom function name
            currentFunction.name = `${currentSnippet.data.namespace}.${currentFunction.name}`;
            customFunctionsMetadata.functions.push(currentFunction);
        });
    });

    //     const snippetBase64OrNull = params.snippetIframesBase64Texts[i];
    //     let $entry = showUI ? $snippetNames.children().eq(i) : null;

    //     if (isNil(snippetBase64OrNull) || snippetBase64OrNull.length === 0) {
    //         if (showUI) {
    //             $entry.addClass(CSS_CLASSES.error);
    //         } else {
    //             heartbeat.messenger.send<LogData>(heartbeat.window, CustomFunctionsMessageType.LOG, {
    //                 timestamp: new Date().getTime(),
    //                 source: 'system',
    //                 type: 'custom functions',
    //                 subtype: 'runner',
    //                 severity: 'error',
    //                 // TODO CUSTOM FUNCTIONS localization
    //                 message: `Could NOT load function "${params.snippetNames[i]}"`
    //             });
    //         }

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

    return customFunctionsMetadata;
}


function handleError(error: Error) {

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
