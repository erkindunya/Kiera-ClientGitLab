//es3 with promise import
/// <reference path ="../node_modules/@types/jquery/index.d.ts" />

// import { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, ODataVersion, ISPHttpClientConfiguration } from '@microsoft/sp-http';

var _spFormDigestRefreshInterval = (<any>window)._spFormDigestRefreshInterval;
var _spPageContextInfo = (<any>window)._spPageContextInfo;
var SP = (<any>window).SP;
var UpdateFormDigest = (<any>window).UpdateFormDigest;

declare var ROOT_SITES: string[];
declare var DEVELOPMENT: boolean;

export class SharePoint {

    private static Prefix: string = 'i:0#.w|';
    private static FilterPrefix: string = "?$";

    private static ajax(options: any): Promise<any> {
        return new Promise(function(resolve, reject) {
            $.ajax(options).done(resolve).fail(reject);
        });
    }

    private static async Get(url: string, prefix: string = "", requiresDigest: boolean = false): Promise<any> {
        let formDigest = requiresDigest ? await this.GetFormDigest(prefix) : null;
        $.support.cors = true;
        return this.ajax({
            url: url,
            method: 'GET',
            crossDomain: true,
            xhrFields: { withCredentials: true },
            headers: {
                'Content-Type': 'application/json; odata=verbose',
                'Accept': 'application/json; odata=verbose',
                'X-RequestDigest': formDigest
            }
        });
    }

    private static async Post(url: string, data: any, prefix: string = ""): Promise<any> {
        let formDigest = await this.GetFormDigest(prefix);
        $.support.cors = true;
        return this.ajax({
            url: url,
            method: 'POST',
            data: JSON.stringify(data),
            crossDomain: true,
            xhrFields: { withCredentials: true },
            headers: {
                'Content-Type': 'application/json; odata=verbose',
                'Accept': 'application/json; odata=verbose',
                'X-RequestDigest': formDigest
            }
        });
    }

    private static async GetFormDigest(prefix: string): Promise<string> {
        if (prefix.startsWith('http')) {
            $.support.cors = true;
            let result = await this.ajax({
                url: prefix + "/_api/contextinfo",
                method: 'POST',
                crossDomain: true,
                xhrFields: { withCredentials: true },
                headers: {
                    'Content-Type': 'application/json; odata=verbose',
                    'Accept': 'application/json; odata=verbose'
                }
            });
            return result.d.GetContextWebInformation.FormDigestValue as string;
        } else {
            UpdateFormDigest(prefix == "" ? "/" : prefix, _spFormDigestRefreshInterval);
            return $("#__REQUESTDIGEST").val();
        }
    }

    private static async Merge(url: string, data: any, prefix: string = ""): Promise<any> {
        let formDigest = await this.GetFormDigest(prefix);
        $.support.cors = true;
        return this.ajax({
            url: url,
            method: 'POST',
            data: JSON.stringify(data),
            crossDomain: true,
            xhrFields: { withCredentials: true },
            headers: {
                'Content-Type': 'application/json; odata=verbose',
                'Accept': 'application/json; odata=verbose',
                'X-RequestDigest': formDigest,
                "X-HTTP-Method": "MERGE",
                "If-Match": "*"
            }
        });
    }

    public static async GetSubSites(prefix: string = '') {
        let sites = [];
        let result = await this.Get(`${prefix}/_api/web/webs`);
        result.d.results.forEach(function (site) {
            sites.push({
                WorkId: 0,
                Title: site.Title,
                Path: site.Url
            });
        });
        return sites;
    }

    public static async GetSites() {
        let sites = [];
        for(let site of ROOT_SITES) {
            try {
                let result = await this.Get(`${site}/_api/search/query?querytext='contentclass:sts_site'&rowlimit=100`, site);
                result.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results.forEach(function (searchItem) {
                    let item = searchItem.Cells.results;
                    sites.push({
                        WorkId: item[2].Value,
                        Title: item[3].Value,
                        Path: item[6].Value
                    });
                });
            } catch(error) {
                console.log(`No access to ${site}`);
            }
        }
        // let result = await this.Get(`/_api/search/query?querytext='contentclass:sts_site'`, '');
        // result.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results.forEach(function (searchItem) {
        //     let item = searchItem.Cells.results;
        //     sites.push({
        //         WorkId: item[2].Value,
        //         Title: item[3].Value,
        //         Path: item[6].Value
        //     });
        // });
        // result = await this.Get(`https://uat-content-mykier/_api/search/query?querytext='contentclass:sts_site'`, 'https://uat-content-mykier');
        // result.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results.forEach(function (searchItem) {
        //     let item = searchItem.Cells.results;
        //     sites.push({
        //         WorkId: item[2].Value,
        //         Title: item[3].Value,
        //         Path: item[6].Value
        //     });
        // });
        return sites;
    }

    public static async GetSiteGroups(prefix: string = '') {
        let result = await this.Get(`${prefix}/_api/web/roleassignments/groups`, prefix);
        return result.d.results;
    }

    public static async AddUserToGroup(groupId: number, loginName: string, prefix: string = '') {
        prefix = prefix == "/" ? "" : prefix;
        return await this.Post(`${prefix}/_api/web/sitegroups(${groupId})/users`, {
            __metadata: { type: 'SP.User' },
            LoginName: loginName
        }, prefix);
    }

    public static async RemoveUserFromGroup(groupId: number, loginName: string, prefix: string = '') {
        return await this.Post(`${prefix}/_api/web/sitegroups(${groupId})/users/removeByLoginName(@v)?@v='${encodeURIComponent(loginName)}'`, {}, prefix);
    }

    public static GetCurrentUserId(): string {
        return _spPageContextInfo.userId;
    }

    public static async GetCurrentUserEmail() {
        let result = await this.Get("/_api/web/currentuser/email");
        return result.d;
    }

    public static async GetListItemByField(listName: string, fieldName: string, fieldValue: string, prefix: string = ''): Promise<any> {
        let result = await this.Get(`${prefix}/_api/web/lists/getbytitle('${listName}')/items?$filter= ${fieldName} eq '${fieldValue}'`, prefix);
        if (result.d != undefined && result.d.results != undefined && result.d.results.length > 0) {
            return result.d.results[0];
        } else {
            return null;
        }
    }

    // see: https://sharepoint.stackexchange.com/questions/129309/how-to-get-permission-of-a-sharepoint-list-for-a-user-using-rest-api
    public static async GetListPermissions(listName: string, prefix: string = ''): Promise<any> {
        try {
            var perms = await this.Get(`${prefix}/_api/web/lists/getbytitle('${listName}')/EffectiveBasePermissions`, prefix);
        } catch (e) {
            return [];
        }

        if (!perms.d || !perms.d.EffectiveBasePermissions) return [];

        var permissions = new SP.BasePermissions();
        permissions.initPropertiesFromJson(perms.d.EffectiveBasePermissions);
        var permLevels = [];
        for (var permLevelName in SP.PermissionKind.prototype) {
            if (SP.PermissionKind.hasOwnProperty(permLevelName)) {
                var permLevel = SP.PermissionKind.parse(permLevelName);
                if (permissions.has(permLevel)) {
                    permLevels.push(permLevelName);
                }
            }
        }
        return permLevels;
    }

    public static async UpdateListItem(listName: string, itemId: string, newItem: any, prefix: string = '') {
        return await this.Merge(`${prefix}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`, newItem, prefix);
    }

    public static async CreateListItem(listName: string, newItem: any, prefix: string = '') {
        return await this.Post(`${prefix}/_api/web/lists/getbytitle('${listName}')/items`, newItem, prefix);
    }

    public static async GetCurrentUser(prefix: string = ''): Promise<any> {
        console.log("Getting current user");
        let result = await this.Get(`${prefix}/_api/web/currentuser?$expand=groups`, prefix);
        return result.d;
    }

    public static async GetUserId(logonName: string, prefix: string = ''): Promise<any>
    {
        let result = await this.Post(`${prefix}/_api/web/ensureUser('${logonName}')`, {}, prefix);
        return result.d.Id;
    }

    // public static async GetUserId(email: string, prefix: string = ''): Promise<number> {
    //     //note: id's are site specific so you will need to provide a prefix if you are having an issue when parsing this to a Person or Group field.
    //     email = Helper.EmailCapitalize(email);
    //     let result = await this.Get(`${prefix}/_api/web/siteusers?$filter=Email eq '${email}'`, prefix);
    //     return result.d.results[0].Id;
    // }

    public static async GetUserLoginName(email: string, prefix: string = ''): Promise<any> {
        // let result = await this.Get(`${prefix}/_api/web/siteusers?$filter=Email eq '${email}'`);
        // let user = email.split('@')[0].toLowerCase();
        // let result = await this.Get(`/_vti_bin/listdata.svc/UserInformationList?$filter=substringof('${user}',tolower(Account))`, prefix);
        let query = email;
        if(!email.toLowerCase().endsWith('@kier.co.uk') && email.includes('@')) {
            query = `fbaMembers:${email}`;
        }
        let result = await this.Post(`/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
            {
                'queryParams':{  
                    '__metadata':{  
                        'type':'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters'  
                    },
                    'MaximumEntitySuggestions':50,  
                    'PrincipalSource':15,  
                    'PrincipalType': 1,  
                    'QueryString': query,
                    'Required':false
                } 
            }
        );
        let parsedResult = JSON.parse(result.d.ClientPeoplePickerSearchUser);
        if(parsedResult.length <= 0) return null;
        if(parsedResult[0].EntityType != 'User') return null;
        // if(parsedResult[0].EntityData.Email.toLowerCase() != email.toLowerCase()) return null;
        return {
            LoginName: parsedResult[0].Key,
            Email: parsedResult[0].EntityData.Email.toLowerCase()
        };
    }

    public static async GetUserGroups(loginName: string, prefix: string = ''): Promise<any> {
        let result = await this.Get(`${prefix}/_api/web/siteusers/?$expand=groups&$filter=LoginName eq '${encodeURIComponent(loginName)}'`);
        if (result.d != undefined && result.d.results != undefined && result.d.results[0].Groups.results.length > 0) {
            return result.d.results[0].Groups.results;
        } else {
            return null;
        }
    }

    public static async GetAllSiteUser(filter: string, prefix: string = ''): Promise<any> {
        return await this.Get(`${prefix}/_api/web/siteusers${this.FilterPrefix}${filter}`, prefix);
    }

    public static async GetSearchItem(searchTerm: string, itemName: string = null) {
        var searchItem = new SearchItem();
        var result: any;
        var items = await this.GetSearchItems(searchTerm);
        if (!itemName) {
            result = items[0];
        } else {
            items.forEach(async item => {
                if (item[3].Value === itemName) {
                    result = item;
                }
            });
        }
        searchItem.Title = result[3].Value;
        searchItem.Url = result[6].Value;
        searchItem.FileType = result[17].Value;
        searchItem.Content = await this.Get(searchItem.Url);
        return searchItem;
    }

    public static async GetSearchItems(searchTerm: string): Promise<SearchItem[]> {
        var searchItems: SearchItem[] = [];
        let result = await this.Get(`/_api/search/query?refinementfilters='fileExtension:equals("aspx")'&querytext='${encodeURIComponent(searchTerm)}'`);
        result.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results.forEach(function (searchItem) {
            searchItems.push(searchItem.Cells.results);
        });
        return searchItems;
    }

    public static async GetWeb(url:string): Promise<string> {
        url = url.replace('/_layouts/15/start.aspx#', '');
        let result = await this.Get(`/_api/sp.web.getweburlfrompageurl(@v)?@v='${decodeURI(url)}'`);
        return result.d.GetWebUrlFromPageUrl;
    }

    public static async GetPageByFullUrl(url: string): Promise<any> {
        let web = await this.GetWeb(url);
        let path = url.replace(/^.*\/\/[^\/]+/, '').split('?')[0];
        return this.GetPageByPath(path, web);
    }

    public static async GetPageByPath(path: string, prefix: string): Promise<any> {
        path = path.replace('/_layouts/15/start.aspx#', '');
        let result = await this.Get(`${prefix}/_api/Web/GetFileByServerRelativeUrl('${decodeURI(path)}')/ListItemAllFields?$expand=ParentList`, prefix);
        return result.d;


        // return null;
    }

    public static async GetListGroups(id: string, prefix: string) {
        prefix = prefix == "/" ? "" : prefix;
        let result = await this.Get(`${prefix}/_api/web/lists/getByID('${id}')/roleassignments/groups`)
        if (result.d.results.length > 0)
            return result.d.results;
        else
            return null;
    }

    public static async GetRoleDefinitions(prefix: string) {
        prefix = prefix == "/" ? "" : prefix;
        let result = await this.Get(`${prefix}/_api/web/roledefinitions`)
        if (result.d.results.length > 0)
            return result.d.results;
        else
            return null;
    }

    public static async CreateSubsite(prefix: string, title: string, url: string, webTemplate: string, useSamePermissionsAsParentSite: boolean = false) {
        prefix = prefix == "/" ? "" : prefix;
        let result = await this.Post(`${prefix}/_api/web/webs/add`, {
            parameters: {
                Title: title,
                Url: url,
                WebTemplate: webTemplate,
                UseSamePermissionsAsParentSite: useSamePermissionsAsParentSite
            }
        }, prefix);

        return result;
    }

    public static GetListItemType(listName: string) {
        if(!DEVELOPMENT && listName == "FBA User Request")
            return "SP.Data.NewFBAUserRequestListItem";
        return "SP.Data." + listName[0].toUpperCase().replace(" ", "_x0020_") + listName.substring(1).split(" ").join("_x0020_") + "ListItem";
    }

    public static async CreateGroup(prefix: string, groupName: string) {
        prefix = prefix == "/" ? "" : prefix;
        let result = await this.Post(`${prefix}/_api/web/sitegroups`, {
            "__metadata": {
                "type": "SP.Group"
            },
            "Title": groupName,
            "Description": "Bot generated group"
        }, prefix);
        return result.d;
    }

    public static async GetParentUrl(url: string)
    {
        return await this.Get(url, '', true);
    }

    public static async AssignRoleToItem(groupId, roleId, prefix, listId, itemId) {
        prefix = prefix == "/" ? "" : prefix;
        let startsWith = `${prefix}/_api/web/lists(guid'${listId}')/items(${itemId})`;
        await this.Post(`${startsWith}/breakroleinheritance(true)`, {}, prefix);
        return await this.Post(`${startsWith}/roleassignments/addroleassignment(principalid=${groupId},roledefid=${roleId})`, {}, prefix);
    }

    public static async AssignRoleToList(groupId, roleId, prefix, listId) {
        prefix = prefix == "/" ? "" : prefix;
        let startsWith = `${prefix}/_api/web/lists(guid'${listId}')`;
        await this.Post(`${startsWith}/breakroleinheritance(true)`, {}, prefix);
        return await this.Post(`${startsWith}/roleassignments/addroleassignment(principalid=${groupId},roledefid=${roleId})`, {}, prefix);
    }

    public static async AssignRoleToSite(groupId, roleId, prefix) {
        prefix = prefix == "/" ? "" : prefix;
        return await this.Post(`${prefix}/_api/web/roleassignments/addroleassignment(principalid=${groupId},roledefid=${roleId})`, {}, prefix);
    }

    // public static async GetUser(email) {
    //     var endpoint = `/_api/web/siteusers?$filter=Email eq '${Helper.EmailCapitalize(email)}'`;
    //     return await this.Get('', endpoint, true);
    // }

    // public static async GetEmail(Id) {
    //     // https://uat-ext.kier.co.uk/_api/web/siteusers?$filter=Email%20eq%20%27Grant.Tapp@kier.co.uk%27
    // }

    public static async GetListFields(list: string, prefix: string, id: number) {
        let endpoint = `${prefix}/_api/web/lists/getbytitle('${list}')/items(${id})`
        let result = await this.Get(endpoint, prefix, true);
        return result.d;
    }

    public static async GetListField(prefix: string, list: string, field: string, id: number) {
        let fields = await this.GetListFields(list, prefix, id);
        return fields[field];
    }

    public static async GetListFieldsAndTypes(prefix) {
        prefix = prefix == "/" ? "" : prefix;
        return await this.Get('', prefix, true);

        // $.ajax({
        //     url: "https://uat-ext.kier.co.uk/sites/KPC/_api/web/lists/getbytitle('DelegateTasks')/items(1)",
        //     method: "GET",
        //     success: data => {
        //         var document = $(data);
        //         var entry = $(document[0].childNodes);
        //         var content = $(entry[0].childNodes[16]);
        //         var properties = $(content[0].childNodes[0].childNodes);

        //         try {
        //             for (var property in properties) {
        //                 if ($(properties[property])[0].nodeType) {
        //                     try {
        //                         // console.log($(properties[property])[0]);
        //                         console.log(`${$(properties[property])[0].localName}: ${$(properties[property])[0].attributes[0].value}`);
        //                     }
        //                     catch (error) { }
        //                 }
        //             }
        //         }
        //         catch (error) { }
        //     }
        // });
    }
}

class Helper {
    public static capitalizeFirstLetter(string): string {
        return string.charAt(0).toUpperCase() + string.slice(1);
    }

    public static EmailCapitalize(email: string): string {
        var splitEmail = email.split('@');
        var splitStart = splitEmail[0].split('.');
        var cappedItems = [];
        splitStart.forEach((item) => {
            cappedItems.push(this.capitalizeFirstLetter(item));
        });
        return cappedItems.join('.') + '@' + splitEmail[1].toLowerCase();
        // var rg = /\W/;
        // var names = email.split(rg);
        // var firstName = this.capitalizeFirstLetter(names[0]);
        // var lastName = this.capitalizeFirstLetter(names[1]);

        // return firstName + "." + lastName + "@kier.co.uk";
    }
}

class SearchItem {
    public Title: string;
    public Url: string;
    public FileType: string;
    public Content: string;
    public ListID: string;
    public SPSiteURL: string;
}

class Error {
    public Success: boolean;
    public Error: string;

    constructor(success: boolean, error: string) {
        this.Success = success;
        this.Error = error;
    }
}