import { storage, environment, post, trustedSnippetManager } from './index';
import { getDisplayLanguage } from '../strings';

export function navigateToRegisterCustomFunctions() {
    const url = environment.current.config.runnerUrl + '/register/custom-functions';

    let allSnippetsToRegisterWithPossibleDuplicate: ICustomFunctionsRegistrationRelevantData[] =
        ([storage.current.lastOpened].concat(storage.snippets.values()))
            .filter(snippet => trustedSnippetManager.isSnippetTrusted(snippet.id, snippet.gist, snippet.gistOwnerId))
            .filter(snippet => snippet.customFunctions && snippet.customFunctions.content && snippet.customFunctions.content.trim().length > 0)
            .map((snippet): ICustomFunctionsRegistrationRelevantData => {
                try {
                    return {
                        name: snippet.name,
                        data: JSON.parse(snippet.customFunctions.content)
                    };
                } catch (e) {
                    throw new Error(`Error parsing metadata for snippet "${snippet.name}`);
                }
            });

    let data: IRegisterCustomFunctionsPostData = {
        snippets: allSnippetsToRegisterWithPossibleDuplicate,
        displayLanguage: getDisplayLanguage()
    };

    return post(url, { data: JSON.stringify(data) });
}
