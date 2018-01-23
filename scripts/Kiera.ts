/// <reference path ="../node_modules/@types/jquery/index.d.ts" />
import { SharePoint } from './SharePoint';
import * as BotChat from 'botframework-webchat';
import FbaEvents from './FbaEvents';
import SearchEvents from './SearchEvents';
import * as CognitiveServices from 'botframework-webchat/CognitiveServices';
import swal from 'sweetalert2';

// webpack provided
declare var DIRECTLINE_SECRET: string;

export class KieraBot {
    botConnection: BotChat.DirectLine;
    speechOptions: BotChat.SpeechOptions;
    botUser: any;
    userId: string;
    public isInitialised: boolean = false;

    constructor() {
        this.botConnection = new BotChat.DirectLine({
            secret: DIRECTLINE_SECRET
        });
    }

    public AddEvents(events: [{ name: string, action: (message: BotChat.EventActivity) => void }]): void {
        events.forEach(event => {
            event.name.split('|').forEach(name => {
                this.AddEvent(name, event.action);
            });
        });
    }

    public AddEvent(name: string, action: (message: BotChat.EventActivity) => void): void {
        this.botConnection.activity$
            .filter((message, index) => {
                if (message.type === "event" && message.name == name)
                    return true;
                return false;
            })
            .subscribe(action.bind(this));
    }

    public SendEvent(name: string, data: any): void {
        this.botConnection
            .postActivity({
                type: "event",
                name: name,
                value: data,
                from: this.botUser
            })
            .subscribe(function (message) {
                // do nothing
            });
    }

    public PreviousCommand(index: number, data: any)
    {
        if(!data)
            data = [];

        return data[index];
    }

    public async InitChat(): Promise<void> {
        var user = await SharePoint.GetCurrentUser();
        var permissionsProcurement = await SharePoint.GetListPermissions('ExternalEmployeeRegistration', "/kiera");
        var permissionsFba = await SharePoint.GetListPermissions('FBA User Request', '/kiera');
        var siteCreation = await SharePoint.GetListPermissions('SiteCollectionCreationList', '/kiera');
        console.log(permissionsProcurement, permissionsFba, siteCreation);
        this.speechOptions = {
            speechRecognizer: new CognitiveServices.SpeechRecognizer({ subscriptionKey: '2c4a1ee3bd624d05893a7a6f04a6dfea' }),
            speechSynthesizer: new CognitiveServices.SpeechSynthesizer({
                gender: CognitiveServices.SynthesisGender.Female,
                subscriptionKey: '2c4a1ee3bd624d05893a7a6f04a6dfea',
                voiceName: 'Microsoft Server Speech Text to Speech Voice (en-US, JessaRUS)'
            })
        }
        this.botUser = { id: user.UserId.NameId, name: user.Title };
        this.userId = user.Id;

        BotChat.App({
            botConnection: this.botConnection,
            user: this.botUser,
            bot: { id: 'KieraBot', name: 'Kiera' },
            speechOptions: this.speechOptions
        }, document.getElementById("bot"));

        $("#aspnetForm").submit(function(event){
            event.preventDefault();
        });

        // help dialog

        $('.help-button').click(function (event) {
            event.preventDefault();
            swal({
              title: 'Kiera Help',
              html: $("#help-button-popup").html(),
              showCloseButton: true,
              grow: 'fullscreen',
              confirmButtonText: 'Close'
            });
        });
        
        //help accordion
        var animTime = 300;
        var clickPolice = false;
        
        $(document).on('touchstart click', '.acc-btn', function(){
            if(!clickPolice){
                clickPolice = true;
                
                var currIndex = $(this).index('.acc-btn'),
                    targetHeight = $('.acc-content-inner').eq(currIndex).outerHeight();
            
                $('.acc-btn h4').removeClass('selected');
                $(this).find('h4').addClass('selected');
                
                $('.acc-content').stop().animate({ height: 0 }, animTime);
                $('.acc-content').eq(currIndex).stop().animate({ height: targetHeight }, animTime);
        
                setTimeout(function(){ clickPolice = false; }, animTime);
            }
        });

        this.SendEvent('welcome', {
            'FBA User Request': permissionsFba,
            'ExternalEmployeeRegistration': permissionsProcurement,
            'SiteCollectionCreationList': siteCreation
        });
        this.isInitialised = true;
    }
}

$(document).ready(function () {
    (<any>window).SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
        let bot = new KieraBot();
        bot.AddEvents(FbaEvents(bot));
        bot.AddEvents(SearchEvents(bot));
        bot.InitChat();
    });

    $('.feedback-button').click(function (event) {
        event.preventDefault();
        swal({
            title: 'Feedback',
            text: 'Any feedback is appreciated and will be used to improve Kiera in the future.',
            input: 'textarea',
            confirmButtonText: 'Send Feedback',
            showCloseButton: true
          })
          .then(feedback => {
            if(!feedback || !feedback.value || feedback.value == "") return false;
            SharePoint.CreateListItem('Kiera Feedback', {
                '__metadata': { 'type': `${SharePoint.GetListItemType('Kiera Feedback')}` },
                'Title': '',
                'Feedback': feedback.value
            }, '/kiera').then((result) => {
                swal('Feedback submitted.');
            });
          });
    });
});