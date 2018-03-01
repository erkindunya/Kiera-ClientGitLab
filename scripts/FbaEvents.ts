import * as BotChat from 'botframework-webchat';
import { SharePoint } from './SharePoint';
import { KieraBot } from './kiera';

declare var MYKIER_URL: string;

function updateUser(user: any): Promise<any> {
	let newItem = $.extend({}, user, {
		'__metadata': { 'type': `${SharePoint.GetListItemType(user.ListName)}` }
	});
	if (newItem.ListName != 'ExternalEmployeeRegistration') {
		newItem.EMail = newItem.Email;
		console.log(newItem);
		delete newItem.Email;
	}
	delete newItem.ListName;
	delete newItem.UrlPrefix;
	return SharePoint.UpdateListItem(user.ListName, user.Id, newItem, user.UrlPrefix);
}

function createUser(user: any): Promise<any> {
	let newItem = $.extend({}, user, {
		'__metadata': { 'type': `${SharePoint.GetListItemType(user.ListName)}` }
	});
	if (newItem.ListName != 'ExternalEmployeeRegistration') {
		newItem.EMail = newItem.Email;
		delete newItem.Email;
	}
	delete newItem.ListName;
	delete newItem.UrlPrefix;
	return SharePoint.CreateListItem(user.ListName, newItem, user.UrlPrefix);
}

function modifyUser(user: any, userId: string, action: string): Promise<any> {
	let newItem = {
		'__metadata': { 'type': `${SharePoint.GetListItemType(user.ListName)}` },
		DeleteFBAUser: action == 'delete' ? 'Delete' : '',
		DisableFBAUser: action == 'disable' ? 'Disable' : (action == 'enable' ? 'Enable' : ''),
		LockFBAUser: action == 'lock' ? 'Lock' : (action == 'unlock' ? 'Unlock' : ''),
		ResetFBAUserPassword: action == 'reset' ? 'Reset' : ''
	};
	return SharePoint.UpdateListItem(user.ListName, user.Id, newItem, user.UrlPrefix);
}

function recordEvent(conversationId: string, content: string, status: string = "Closed") {
	SharePoint.GetCurrentUserEmail().then(function (response) {
		let item = {
			"__metadata": {
				"type": SharePoint.GetListItemType("Support Request")
			},
			"ConversationId": conversationId,
			"Content": content,
			"Title": response.Email,
			"Status": status
		};
		SharePoint.CreateListItem('Support Request', item, '/kiera').then(function (response) {
			console.log("Action recorded.");
		}).catch(function (error) {
			console.log("Action failed to record");
		});
	}).catch(function (error) {
		console.log("Recording event failed.");
	});
}

let FbaEvents: (kiera: KieraBot) => { name: string, action: (message: BotChat.EventActivity) => void }[] = function (kiera: KieraBot) {
	return [
		{
			name: 'getfbauser',
			action: (message) => {
				let email = message.value.Email;
				let listName = message.value.ListName;
				let urlPrefix = message.value.UrlPrefix;
				SharePoint.GetListItemByField(listName, listName == 'ExternalEmployeeRegistration' ? 'Email' : 'EMail', email, urlPrefix).then(result => {
					if (result) {
						if (listName != 'ExternalEmployeeRegistration') {
							result.Email = result.EMail;
							delete result.EMail;
						}
						kiera.SendEvent(listName == 'ExternalEmployeeRegistration' ? 'setprocurementuser' : 'setfbauser', result);
					}
					else
						kiera.SendEvent('nouserfound', message.value.Email);
				}).catch(error => {
					kiera.SendEvent('error', error);
				});
			}
		},
		{
			name: 'updatefbauser',
			action: (message) => {
				let email = message.value.OldEmail;
				let listName = message.value.ListName;
				let urlPrefix = message.value.UrlPrefix;
				SharePoint.GetListItemByField(listName, listName == 'ExternalEmployeeRegistration' ? 'Email' : 'EMail', email, urlPrefix).then(user => {
					if (user) {
						if (message.value.OldEmail != message.value.Email) {
							SharePoint.GetListItemByField(listName, listName == 'ExternalEmployeeRegistration' ? 'Email' : 'EMail', message.value.Email, urlPrefix).then(user => {
								if (!user) {
									// delete message.value.OldEmail;
									updateUser(message.value).then(result => {
										kiera.SendEvent('updatedfbauser', message.value.Email);
										recordEvent(message.conversation.id, "Updated FBA User");
									}).catch(error => {
										kiera.SendEvent('error', error);
									});
								} else {
									kiera.SendEvent('useralreadyexists', message.value.Email);
								}
							}).catch(error => {
								kiera.SendEvent('error', error);
							});
						} else {
							// delete message.value.OldEmail;
							updateUser(message.value).then(result => {
								kiera.SendEvent('updatedfbauser', message.value.Email);
							}).catch(error => {
								kiera.SendEvent('error', error);
							});
						}
					} else {
						kiera.SendEvent('nouserfound', message.value.OldEmail);
					}
				}).catch(error => {
					kiera.SendEvent('error', error);
				});
			}
		},
		{
			name: 'createfbauser',
			action: (message) => {
				let listName = message.value.ListName;
				let urlPrefix = message.value.UrlPrefix;
				let email = message.value.Email;
				SharePoint.GetListItemByField(listName, listName == 'ExternalEmployeeRegistration' ? 'Email' : 'EMail', email, urlPrefix).then(user => {
					if (!user) {
						createUser(message.value).then(result => {
							kiera.SendEvent('createdfbauser', message.value.Email);
							recordEvent(message.conversation.id, "Created FBA User");
						}).catch(error => {
							kiera.SendEvent('error', error);
						});
					} else {
						kiera.SendEvent('useralreadyexists', message.value.Email);
					}
				}).catch(error => {
					kiera.SendEvent('error', error);
				});
			}
		},

		{
			name: 'deletefbauser|lockfbauser|resetfbauser|unlockfbauser|disablefbauser|enablefbauser',
			action: (message) => {
				let listName = message.value.ListName;
				let urlPrefix = message.value.UrlPrefix;
				let email = message.value.Email;
				let actionName = message.name.replace('fbauser', '');
				SharePoint.GetListItemByField(listName, listName == 'ExternalEmployeeRegistration' ? 'Email' : 'EMail', email, urlPrefix).then(user => {
					if (user) {
						user.ListName = listName;
						user.UrlPrefix = urlPrefix;
						modifyUser(user, kiera.userId, actionName).then(result => {
							kiera.SendEvent(message.name + 'done', message.value.Email);
							recordEvent(message.conversation.id, `${actionName} FBA User`);
						}).catch(error => {
							kiera.SendEvent('error', error);
						});
					} else {
						kiera.SendEvent('nouserfound', message.value.Email);
					}
				}).catch(error => {
					kiera.SendEvent('error', error);
				});
			}
		},
		{
			name: 'getsitecollections|getsubsite',
			action: async (message) => {
				let actionName = message.name.replace('get', 'set');
				let teamName = message.value.TeamName;
				SharePoint.GetSites().then(sites => {
					if (sites) {
						kiera.SendEvent(actionName, {
							Sites: sites,
							TeamName: teamName
						});
					} else {
						kiera.SendEvent('nositesfound', '');
					}
				}).catch(error => {
					kiera.SendEvent('error', error);
				});
			}
		},
		{
			name: 'getsites|getusersites',
			action: async (message) => {
				let email: string = message.value.Email;
				let actionName = message.name.replace('get', 'set');
				SharePoint.GetUserLoginName(email).then((loginName) => {
					if (!loginName) {
						kiera.SendEvent('nouserfound', email);
					} else if (loginName.Email === email.toLowerCase()) {
						SharePoint.GetSites().then(sites => {
							if (sites) {
								kiera.SendEvent(actionName, {
									LoginName: loginName.LoginName,
									Sites: sites,
									Email: loginName.Email
								});
							} else {
								kiera.SendEvent('nositesfound', '');
							}
						}).catch(error => {
							kiera.SendEvent('error', error);
						});
					} else {
						// check if email correct, and if so, return to action passing email merged with state
						kiera.SendEvent('confirmuser', {
							LoginName: loginName.LoginName,
							Email: loginName.Email,
							ActionName: message.name,
							State: {
								TeamName: message.value.TeamName
							}
						});
					}
				}).catch(error => {
					kiera.SendEvent('error', error);
				});
			}
		},
		{
			name: 'getsubsites',
			action: async (message) => {
				let loginName: string = message.value.LoginName;
				let teamName: string = message.value.TeamName;
				let urlPrefix: string = message.value.UrlPrefix;
				SharePoint.GetSubSites(urlPrefix).then(sites => {
					if (sites && sites.length > 0) {
						if(!loginName){
							kiera.SendEvent("setsubsite", {
								Sites: sites,
								UrlPrefix: urlPrefix,
								TeamName: teamName
							});
						}
						kiera.SendEvent("setsites", {
							LoginName: loginName,
							Sites: sites
						});
					} else {
						kiera.SendEvent('nosubsites', {
							LoginName: loginName,
							UrlPrefix: urlPrefix
						});
					}
				}).catch(error => {
					kiera.SendEvent('error', error);
				});
			}
		},
		{
			name: 'getsitegroups',
			action: (message) => {
				let loginName = message.value.LoginName;
				let urlPrefix = message.value.UrlPrefix;
				SharePoint.GetSiteGroups(urlPrefix).then(groups => {
					if (groups) {
						kiera.SendEvent('setgroups', {
							LoginName: loginName,
							Groups: groups,
							UrlPrefix: urlPrefix
						});
					} else {
						kiera.SendEvent('nogroupsfound', urlPrefix);
					}
				}).catch(error => {
					kiera.SendEvent('error', error);
				});
			}
		},
		{
			name: 'getusergroups',
			action: (message) => {
				let loginName = message.value.LoginName;
				let urlPrefix = message.value.UrlPrefix;
				SharePoint.GetUserGroups(loginName, urlPrefix).then(groups => {
					if (groups) {
						kiera.SendEvent('setusergroups', {
							LoginName: loginName,
							Groups: groups,
							UrlPrefix: urlPrefix
						});
					} else {
						kiera.SendEvent('nogroupsfound', loginName);
					}
				}).catch(error => {
					kiera.SendEvent('error', error);
				});
			}
		},
		{
			name: 'addusergroup',
			action: (message) => {
				let urlPrefix = message.value.UrlPrefix;
				let loginName = message.value.LoginName;
				let groupId = message.value.GroupId;
				SharePoint.AddUserToGroup(groupId, loginName, urlPrefix).then((result) => {
					if (result) {
						kiera.SendEvent('addedusergroup', groupId);
						recordEvent(message.conversation.id, `Added User to Group`);
					} else {
						kiera.SendEvent('addtogroupfailed', groupId);
					}
				}).catch(error => {
					if (error.status == 401)
						kiera.SendEvent('permissiondenied', `Permission denied when adding ${loginName} to group ${groupId} in ${urlPrefix}`);
					else
						kiera.SendEvent('error', error);
				});
			}
		},
		{
			name: 'removeusergroup',
			action: (message) => {
				let urlPrefix = message.value.UrlPrefix;
				let loginName = message.value.LoginName;
				let groupId = message.value.GroupId;
				SharePoint.RemoveUserFromGroup(groupId, loginName, urlPrefix).then((result) => {
					if (result) {
						kiera.SendEvent('removedusergroup', groupId);
						recordEvent(message.conversation.id, `Removed User From Group`);
					} else {
						kiera.SendEvent('removefromgroupfailed', groupId);
					}
				}).catch(error => {
					kiera.SendEvent('error', error);
				});
			}
		},
		{
			name: 'geturlgroups',
			action: async (message) => {
				let fullUrl = message.value.Path;
				fullUrl = fullUrl.replace('http://', 'https://');
				let path = fullUrl.replace(/^.*\/\/[^\/]+/, '').split('?')[0];
				// let prefix = path.startsWith('/sites') ? path.split("/").slice(0, 3).join("/") : "";
				// console.log(path);
				// console.log(prefix);
				// let path = message.value.Path;
				let email = message.value.Email;
				try {
					let user = await SharePoint.GetUserLoginName(email);
					if (!user) {
						kiera.SendEvent('nouserfound', email);
					} else if (user.Email === email.toLowerCase()) {
						let loginName = user.LoginName;
						let prefix = await SharePoint.GetWeb(fullUrl);
						let isSite = false;
						if(prefix.toLowerCase().endsWith(path.toLowerCase())) isSite = true;

						let result = null;
						if(!isSite) {
							result = await SharePoint.GetPageByPath(path, prefix);
							if (!result || !result.ListItemAllFields) isSite = true;
						}

						// if it ends with the same then its not a list item but a site
						if(isSite) {
							SharePoint.GetSiteGroups(prefix).then(groups => {
								if (groups) {
									kiera.SendEvent('setgroups', {
										LoginName: loginName,
										Groups: groups,
										UrlPrefix: prefix
									});
								} else {
									kiera.SendEvent('nogroupsfound', prefix);
								}
							}).catch(error => {
								kiera.SendEvent('error', error);
							});
						} else { // otherwise its a list item
							kiera.SendEvent('pagefound', path);
							let groups = await SharePoint.GetListGroups(result.ParentList.Id, prefix);
							if (groups) {
								kiera.SendEvent('setgroups', {
									LoginName: loginName,
									Groups: groups,
									ItemId: result.Id,
									ListId: result.ParentList.Id,
									UrlPrefix: result.ParentList.ParentWebUrl
								});
							} else {
								kiera.SendEvent('nogroupsfound', loginName);
							}
						}
					} else {
						// check if email correct, and if so, return to action passing email merged with state
						kiera.SendEvent('confirmuser', {
							LoginName: user.LoginName,
							Email: user.Email,
							ActionName: message.name,
							State: {
								Path: message.value.Path
							}
						});
					}
				} catch (error) {
					kiera.SendEvent('error', error);
				}
			}
		},
		{
			name: 'getavailablepermissions',
			action: async (message) => {
				let loginName = message.value.LoginName;
				let urlPrefix = message.value.UrlPrefix;
				let listId = message.value.ListId;
				let itemId = message.value.ItemId;
				try {
					var roleAssignments = await SharePoint.GetRoleDefinitions(urlPrefix);
					if (roleAssignments) {
						kiera.SendEvent('setroles', {
							LoginName: loginName,
							UrlPrefix: urlPrefix,
							ListId: listId,
							ItemId: itemId,
							Roles: roleAssignments
						});
					} else {
						kiera.SendEvent('norolesfound', urlPrefix);
					}
				} catch (error) {
					kiera.SendEvent('error', error);
				}

			}
		},
		{
			name: 'creategroup',
			action: async (message) => {
				let loginName = message.value.LoginName;
				let urlPrefix = message.value.UrlPrefix;
				let listId = message.value.ListId;
				let itemId = message.value.ItemId;
				let roleId = message.value.RoleId;
				let groupName = message.value.GroupName;
				try {
					let group = await SharePoint.CreateGroup(urlPrefix, groupName);
					if (group) {
						if (itemId)
							await SharePoint.AssignRoleToItem(group.Id, roleId, urlPrefix, listId, itemId);
						else if (listId)
							await SharePoint.AssignRoleToList(group.Id, roleId, urlPrefix, listId);
						else
							await SharePoint.AssignRoleToSite(group.Id, roleId, urlPrefix);
						kiera.SendEvent('permissionsassigned', groupName);

						await SharePoint.AddUserToGroup(group.Id, loginName, urlPrefix);
						kiera.SendEvent('addedusergroup', group.Id);

						recordEvent(message.conversation.id, `Created Group and Assigned User`);
					} else {
						kiera.SendEvent('creategroupfailed', groupName);
					}
				} catch (error) {
					kiera.SendEvent('error', error);
				}
			}
		},
		{
			name: 'createbluefinsite',
			action: async (message) => {
				try {
					let project = {
						"__metadata": {
							"type": SharePoint.GetListItemType('projects')
						},
						"Business_x0020_Unit": message.value.Business_x0020_Unit,
						"Opportunity_x0020_number": message.value.Opportunity_x0020_number,
						"ProjectNumber": message.value.ProjectNumber,
						"Tender_x0020_Number": message.value.Tender_x0020_Number,
						"Title": message.value.Title
					};

					let createdProject = await SharePoint.CreateListItem('projects', project, '/sites/projects');

					if(createdProject)
					{
						kiera.SendEvent('createdbluefinsite', project.ProjectNumber);
						recordEvent(message.conversation.id, 'Created Bluefin site');
					}
					else
					{
						kiera.SendEvent('failedbluefinsite', project.ProjectNumber);
						recordEvent(message.conversation.id, 'Failure to create bluefin site', 'Open');
					}
				} catch (error) {
					kiera.SendEvent('error', error);
				}
			}
		},
		{
			name: 'createsubsite',
			action: async (message) => {
				try {
					//templates: STS#0 (Team Site) 
					let urlPrefix = message.value.UrlPrefix;
					let teamName = message.value.TeamName;
					let template = 'STS#0';

					try
					{
						let site = await SharePoint.CreateSubsite(urlPrefix, teamName, teamName, template);
						let parentUrl = await SharePoint.GetParentUrl(site.d.ParentWeb.__deferred.uri);
						let ownerId = await SharePoint.CreateGroup(urlPrefix, `${teamName} Owners`);
						let visitorId = await SharePoint.CreateGroup(urlPrefix, `${teamName} Visitors`);
						let memberId = await SharePoint.CreateGroup(urlPrefix, `${teamName} Members`);
	
						await SharePoint.AssignRoleToSite(ownerId.Id, '1073741829', site.d.Url);
						await SharePoint.AssignRoleToSite(visitorId.Id, '1073741924', site.d.Url);
						await SharePoint.AssignRoleToSite(memberId.Id, '1073741827', site.d.Url);
	
						if (site) {
							kiera.SendEvent('createdteamsite', site.d.Url);
							recordEvent(message.conversation.id, `Created Team Site`);
						}
						else {
							kiera.SendEvent('failedteamsite', teamName);
						}
					}
					catch(error)
					{

						kiera.SendEvent('sitealreadyexists', teamName);
					}
				}
				catch (error) {
					kiera.SendEvent('error', error);
				}
			}
		},
		
		{
			name: 'createmykier',
			action: async (message) => {
				try {
					//templates: STS#0 (Team Site) 
					let urlPrefix = MYKIER_URL;
					let teamName = message.value.SiteName;
					let template = 'CMSPUBLISHING#0';

					try
					{
						// dont create groups for my kier
						let site = await SharePoint.CreateSubsite(urlPrefix, teamName, teamName, template, true);
	
						if (site) {
							kiera.SendEvent('createdteamsite', site.d.Url);
							recordEvent(message.conversation.id, `Created MyKier Site`);
						}
						else {
							kiera.SendEvent('failedteamsite', teamName);
						}
					}
					catch(error)
					{

						kiera.SendEvent('sitealreadyexists', teamName);
					}
				}
				catch (error) {
					kiera.SendEvent('error', error);
				}
			}
		},
		{
			name: 'getharmonieaccount',
			action: async (message) => {
				try {
					let machineName = message.value.MachineName;
					let usersName = message.value.UsersName;
					let item = {
						"__metadata": {
							"type": SharePoint.GetListItemType("Harmonie")
						},
						"Machine_x0020_Name": machineName,
						"Users_x0020_Name": usersName,
						"Title": usersName
					};
					await SharePoint.CreateListItem('Harmon.ie', item, "/kiera/");
					kiera.SendEvent('createdharmonieaccount', machineName);
					recordEvent(message.conversation.id, `Account Added to Harmon.ie AD Group`);
				}
				catch (error) {
					kiera.SendEvent('error', error);
				}
			}
		},
		{
			name: 'startptpworkflow',
			action: async (message) => {
                try {
                    let data = {
                        "__metadata": {
                            "type": SharePoint.GetListItemType('projects')
                        },
                        "StartWorkflow": message.value.WorkFlowName
                    };
                    await SharePoint.UpdateListItem('projects', message.value.ProjectId, data, '/sites/KPC');
                    kiera.SendEvent('startworkflow', message.value.WorkFlowName);
                    recordEvent(message.conversation.id, `Started PTP Workflow`);
                }
                catch (error) {
                    kiera.SendEvent('error', error)
                }
            }
		},
		{
			name: 'createptpproject',
			action: async (message) => {
				try {
					let ids = [];

					for (var email of message.value.ProjectManagerId) {
						let userid = await SharePoint.GetUserId(email, '/sites/kpc');
						ids.push(userid);
					};

					let project = {
						"__metadata": {
							"type": SharePoint.GetListItemType('Projects')
						},
						"Title": message.value.Title,
						"CategoryDescription": message.value.CategoryDescription,
						"TotalValuetoKier": message.value.TotalValuetoKier,
						"ProjectManagerId": { results: ids },
						"Kier_x0020_BU": message.value.Kier_x0020_BU,
						"Kier_x0020_Division": message.value.Kier_x0020_Divison,
						"NextApprovalFromId": ""
					};

					let approver = await SharePoint.GetListItemByField('ApproversDetails', 'Title', message.value.Kier_x0020_BU, '/sites/KPC');
					let approverid = approver.ApproversId.results[0];
					project.NextApprovalFromId = approverid;

					await SharePoint.CreateListItem('Projects', project, '/sites/KPC');
					kiera.SendEvent('createdptpaccount', project.Title);

					recordEvent(message.conversation.id, `Created PTP Project`);
				}
				catch (error) {
					kiera.SendEvent('error', error);
				}
			}
		},
		{
			name: 'sendsupportrequest|logexitrequest',
			action: async (message) => {
				let content = message.value;
				try {
					let contentObj = JSON.parse(content);
					content = contentObj.responseJSON.error.message.value;
				}
				catch (error) {
					// leave content
				}

				try {
					let conversationId = message.conversation.id;
					let item = {
						"__metadata": {
							"type": SharePoint.GetListItemType("Support Request")
						},
						"ConversationId": conversationId,
						"Content": content,
						"Title": (await SharePoint.GetCurrentUserEmail()).Email,
						"Status": message.name == 'sendsupportrequest' ? "Open" : "Closed"
					};
					await SharePoint.CreateListItem('Support Request', item, '/kiera');
					kiera.SendEvent('supportrequestsent', message.name);
				}
				catch (error) {
					kiera.SendEvent('error', error);
				}
			}
		},
		{
			name: 'createdelegation',
			action: async (message) => {
				try {
					let userEmail = await SharePoint.GetCurrentUserEmail();
					let delegation = {
						"__metadata": {
							"type": SharePoint.GetListItemType('DelegateTasks')
						},
						"DelegateToId": await SharePoint.GetUserId(message.value.Name),
						"DelegateFromId": SharePoint.GetCurrentUserId(),
						"StartDate": message.value.StartDate,
						"EndDate": message.value.EndDate,
						"CancelDe": message.value.ToCancel = true ? 1 : 0,
						"Title": `${userEmail.Email} to ${message.value.Name}`
					};
					await SharePoint.CreateListItem('DelegateTasks', delegation, '/sites/KPC');
					kiera.SendEvent('createddelegation', delegation.Title);
					recordEvent(message.conversation.id, `Created PTP Delegation`);
				}
				catch (error) {
					kiera.SendEvent('error', error);
				}
			}
		},
		{
			name: 'restartworkflow',
			action: async (message) => {
				try {
					let data = {
						"__metadata": {
							"type": SharePoint.GetListItemType('projects')
						},
						"RestartWorkflow": "Restart"
					};
					await SharePoint.UpdateListItem('projects', message.value.ProjectId, data, '/sites/KPC');
					kiera.SendEvent('restartedworkflow', null);
					recordEvent(message.conversation.id, `Restarted PTP Workflow`);
				}
				catch (error) {
					kiera.SendEvent('error', error)
				}
			}
		},
		{
			name: 'stopworkflow',
			action: async (message) => {
				try {
					let data = {
						"__metadata": {
							"type": SharePoint.GetListItemType('projects')
						},
						"RestartWorkflow": "Stop"
					};
					await SharePoint.UpdateListItem('projects', message.value.ProjectId, data, '/sites/KPC');
					kiera.SendEvent('stoppedworkflow', null);
					recordEvent(message.conversation.id, `Stopped PTP Workflow`);
				}
				catch (error) {
					kiera.SendEvent('error', error)
				}
			}
		},
		{
			name: 'createquery',
			action: async (message) => {
				try {
					let column = message.value.Column;
					let id = message.value.ID;
					let columnTitle = message.value.ColumnTitle;
					let result = await SharePoint.GetListFields('Projects', '/sites/KPC', message.value.ID);
					let field = result[message.value.Column];

					let title = result.Title;

					if (!field)
						kiera.SendEvent('nocolumn', {
							Column: column,
							ID: id,
							Title: title,
							ColumnTitle: columnTitle
						});
					else if (field.results)
						kiera.SendEvent('ptpquery', {
							Column: column,
							ColumnTitle: columnTitle,
							Result: field.results[0],
							Title: title,
							ID: id
						});
					else
						kiera.SendEvent('ptpquery', {
							Column: column,
							ColumnTitle: columnTitle,
							Result: field,
							Title: title,
							ID: id
						});
				}
				catch (error) {
					// console.log(error);
					if (error.status == 404)
						kiera.SendEvent('noprojectfound', message.value.ID);
					else
						kiera.SendEvent('error', error);
				}
			}
		},
		{
			name: 'setptpvalue',
			action: async (message) => {
				try {
					let column = message.value.Column;
					let id = message.value.ID;
					let newValue = message.value.NewValue;

					let newItem = {
						'__metadata': { 'type': `${SharePoint.GetListItemType('Projects')}` }
					};
					newItem[column] = newValue;
					let result = await SharePoint.UpdateListItem('Projects', id, newItem, '/sites/KPC');
					kiera.SendEvent('ptpvalueupdated', '');
					recordEvent(message.conversation.id, `Updated PTP List ${id}, Set '${column}' to ${newValue}`);
				}
				catch (error) {
					// console.log(error);
					if (error.status == 404)
						kiera.SendEvent('noprojectfound', message.value.ID);
					else
						kiera.SendEvent('error', error);
				}
			}
		}
	]
};

export default FbaEvents;