import { endianness } from "os";

class LuisHelper {
    public static Get(url: string): Promise<any> {
        return new Promise((resolve, reject) => {
            $.ajax({
                url: url,
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json',
                    'Ocp-Apim-Subscription-Key': '5e689a68fe1c4a58a4da39ba62b3f4d9'
                },
                success: function (data) {
                    resolve(data);
                },
                error: function (error) {
                    reject(error.responseText);
                }
            });
        });
    }

    public static Post(url: string, data: any): Promise<any> {
        return new Promise((resolve, reject) => {
            $.ajax({
                url: url,
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json',
                    'Ocp-Apim-Subscription-Key': '5e689a68fe1c4a58a4da39ba62b3f4d9'
                },
                data: data,
                success: function (data) {
                    resolve(data);
                },
                error: function (error) {
                    reject(error.responseText);
                }
            });
        });
    }

    public static async AddLuisIntent(intentId: string, intent: string) {
        let endpoint = `https://eu.luis.ai/applications/2804a910-19a6-43f2-ab81-3c9fba765bb6/versions/0.1/build/intents/${intentId}`;
        await this.Post(endpoint, intent);
    }

    public static async AddLuisIntents(intentId: string, intents: string[]) {
        let endpoint = `https://eu.luis.ai/applications/2804a910-19a6-43f2-ab81-3c9fba765bb6/versions/0.1/build/intents/${intentId}`;
        await this.Post(endpoint, intents);
    }

    public static async GetLuisIntents(): Promise<Intent[]> {
        let endpoint = 'https://westeurope.api.cognitive.microsoft.com/luis/api/v2.0/apps/2804a910-19a6-43f2-ab81-3c9fba765bb6/versions/0.1/intents?take=500';
        let intents = await this.Get(endpoint);
        return intents;
    }

    public static async GetLuisIntent(intentName: string): Promise<any> {
        let intents: Intent[] = await this.GetLuisIntents();
        let intent = new Intent();
        let result: any;
        intents.forEach(async item => {
            if (item.name == intentName) {
                result = item;
            }
        });

        intent.name = result.name;
        intent.id = result.id;

        return intent;
    }
}

class Intent {
    id: string;
    name: string;
}