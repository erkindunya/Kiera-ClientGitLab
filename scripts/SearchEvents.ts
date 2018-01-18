
import * as BotChat from 'botframework-webchat';
import { SharePoint } from './SharePoint';
import {KieraBot} from './kiera';

let SearchEvents : (kiera: KieraBot) => [{name: string, action: (message : BotChat.EventActivity) => void}] = function (kiera: KieraBot) {
    return [
    {
        name: 'search',
        action: (message) => {
            let query = message.value;
            SharePoint.GetSearchItem(query).then(result => {
                if(result)
                    kiera.SendEvent('searchresult', result);
                else
                    kiera.SendEvent('noresults', message.value.email);
            }).catch(error => {
                kiera.SendEvent('error', error);
            });
        }
    }
]
};

export default SearchEvents;