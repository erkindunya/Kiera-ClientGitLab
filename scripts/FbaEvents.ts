import * as BotChat from 'botframework-webchat';
import { SharePoint } from './SharePoint';
import { KieraBot } from './kiera';

function updateUser(user: any): Promise<any> {
	let newItem = $.extend({}, user, {
		'__metadata': { 'type': `${SharePoint.GetListItemType(user.ListName)}` }
	});
	if(newItem.ListName != 'ExternalEmployeeRegistration') {
		newItem.EMail = newItem.Email;
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
	if(newItem.ListName != 'ExternalEmployeeRegistration') {
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

let FbaEvents: (kiera: KieraBot) => [{ name: string, action: (message: BotChat.EventActivity) => void }] = function (kiera: KieraBot) {
	return [
		{
			name: 'getfbauser',
			action: (message) => {
				let email = message.value.Email;
				let listName = message.value.ListName;
				let urlPrefix = message.value.UrlPrefix;
				SharePoint.GetListItemByField(listName, listName == 'ExternalEmployeeRegistration' ? 'Email' : 'EMail', email, urlPrefix).then(result => {
					if (result){
						if(listName != 'ExternalEmployeeRegistration') {
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
									delete message.value.OldEmail;
									updateUser(message.value).then(result => {
										kiera.SendEvent('updatedfbauser', message.value.Email);
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
							delete message.value.OldEmail;
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
			name: 'getsites|getusersites|getsitecollections|getsubsite',
			action: async (message) => {
				let currentEmail = await SharePoint.GetCurrentUserEmail();
				currentEmail = currentEmail.Email;
				let email = message.value.Email || currentEmail;
				let actionName = message.name.replace('get', 'set');
				let teamName = message.value.TeamName;
				SharePoint.GetUserLoginName(email).then((loginName) => {
					if (loginName) {
						SharePoint.GetSubSites().then(sites => {
							if (sites) {
								kiera.SendEvent(actionName, {
									LoginName: loginName,
									Sites: sites,
									TeamName: teamName
								});
							} else {
								kiera.SendEvent('nositesfound', '');
							}
						}).catch(error => {
							kiera.SendEvent('error', error);
						});
					} else {
						kiera.SendEvent('nouserfound', email);
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
				let path = message.value.Path.replace(/^.*\/\/[^\/]+/, '').split('?')[0];
				// let prefix = path.startsWith('/sites') ? path.split("/").slice(0, 3).join("/") : "";
				// console.log(path);
				// console.log(prefix);
				// let path = message.value.Path;
				let email = message.value.Email;
				try {
					let loginName = await SharePoint.GetUserLoginName(email)
					if (loginName) {
						let prefix = await SharePoint.GetWeb(fullUrl);
						let result = await SharePoint.GetPageByPath(path, prefix);
						if (result) {
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
						} else {
							kiera.SendEvent('nopagefound', path);
						}
					} else {
						kiera.SendEvent('nouserfound', email);
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
					} else {
						kiera.SendEvent('creategroupfailed', groupName);
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
					SharePoint.CreateSubsite(urlPrefix, teamName, teamName, template);
					kiera.SendEvent('createdteamsite', teamName);
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
					await SharePoint.CreateListItem('Harmonie', item);
					kiera.SendEvent('createdharmonieaccount', machineName);
				}
				catch (error) {
					kiera.SendEvent('error', error);
				}
			}
		},
		{
			name: 'createptpuser',
			action: async (message) => {
				try {
					let ids = [];

					for (var email in message.value.ProjectManagerId) {
						let userid = await SharePoint.GetUserId(message.value.ProjectManagerId[email]);
						ids.push(userid);
					};

					let project = {
						"__metadata": {
							"type": SharePoint.GetListItemType('Projects')
						},
						"Title": message.value.Title,
						"CategoryDescription": message.value.CategoryDescription,
						// "approved": message.value.Approved,
						"TotalValuetoKier": message.value.TotalValuetoKier,
						"ProjectManagerId": { results: ids },
						"Kier_x0020_BU": message.value.Kier_x0020_BU,
						"Kier_x0020_Division": message.value.Kier_x0020_Divison
					};

					SharePoint.CreateListItem('Projects', project, '/sites/KPC');
					kiera.SendEvent('createdptpaccount', project.Title);
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
				catch (error){
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
						"Title": `${userEmail.Email} to ${userEmail.Email}`
					};
					await SharePoint.CreateListItem('DelegateTasks', delegation, '/sites/KPC');
					kiera.SendEvent('createddelegation', delegation.Title);
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
							"type": SharePoint.GetListItemType('PtpRestart')
						},
						"Title": message.value.Name
					};
					await SharePoint.CreateListItem('PtpRestart', data, '/sites/KPC');
					kiera.SendEvent('restartedworkflow', null);
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
					let data = {
						"fieldData": await SharePoint.GetListField('/sites/KPC', 'Projects',  message.value.Column, message.value.ID)
					};

					if(data.fieldData)
						kiera.SendEvent("ptpquery", data.fieldData);
					else
						kiera.SendEvent("nocolumn", message.value);
				}
				catch (error) {
					kiera.SendEvent('error', error);
					console.log(error);
					console.log(message);
				}
			}
		}
	]
};

export default FbaEvents;