// Copyright (C) 2020 Potix Corporation. All Rights Reserved
// History: 2020/12/22 5:55 PM
// Author: charlie<charliehsieh@potix.com>

const {
  TurnContext,
  MessageFactory,
  TeamsActivityHandler,
  CardFactory
} = require('botbuilder');
const { CardTemplates } = require('../model/cardtemplates');
const { TeamsHttp } = require('../utils/teamsHttp');
const { getUserToken, isUserLogin } = require('../utils/tokenManager');
const { getConversationId } = require('../utils/utils');
const utils = require('../utils/utils');
const domainName = process.env.DomainName;

class BotActivityHandler extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      TurnContext.removeRecipientMention(context.activity);
      const command = new String(context.activity.text).trim().toLocaleLowerCase();
      const isLogin = await isUserLogin(context.activity.from.id);

      switch (command) {
        /*
         * only for welcome card test, delete this before release
         */
        case 'welcome':
          const welcomeCard = CardFactory.adaptiveCard(CardTemplates.welcomeCard());
          await context.sendActivity(MessageFactory.attachment(welcomeCard));
          break;

        case 'add task':
        case 'create task': {
          let respondCard;
          if (isLogin)
            respondCard = CardFactory.adaptiveCard(CardTemplates.addTaskButton());
          else
            respondCard = CardFactory.adaptiveCard(CardTemplates.needToLoginCard('adding a new task'));

          await context.sendActivity(MessageFactory.attachment(respondCard));
          break;
        }
        case 'link project': {
          let respondCard;
          if (isLogin)
            respondCard = CardFactory.adaptiveCard(CardTemplates.linkProjectButton());
          else
            respondCard = CardFactory.adaptiveCard(CardTemplates.needToLoginCard('linking a project'));

          await context.sendActivity(MessageFactory.attachment(respondCard));
          break;
        }
        case 'follow project': {
          let respondCard;
          if (isLogin)
            respondCard = CardFactory.adaptiveCard(CardTemplates.followProjectButton());
          else
            respondCard = CardFactory.adaptiveCard(CardTemplates.needToLoginCard('following a project'));

          await context.sendActivity(MessageFactory.attachment(respondCard));
          break;
        }
        case 'login': {
          const conversationType = context.activity.conversation.conversationType;
          const conversationRef = TurnContext.getConversationReference(context.activity);
          if (conversationType === 'groupChat' || conversationType === 'channel') {
            await context.sendActivity(`Thanks ${context.activity.from.name}, I've sent you a direct message to help you do this. If you don't see the message, try adding the Asana app first`);
          }

          let returnMessage;
          if (isLogin) {
            returnMessage = MessageFactory.text('Hey, you’re already logged in.');
          } else {
            const loginButton = CardFactory.adaptiveCard(CardTemplates.loginButton());
            returnMessage = MessageFactory.attachment(loginButton);
          }
          context.adapter.createConversation(conversationRef, async context => {
            await context.sendActivity(returnMessage);
          });
          break;
        }
        case 'logout': {
          const conversationType = context.activity.conversation.conversationType;
          if (conversationType === 'groupChat' || conversationType === 'channel') {
            await context.sendActivity(`Thanks ${context.activity.from.name}, I've sent you a direct message to help you do this. If you don't see the message, try adding the Asana app first`);
            const conversationRef = TurnContext.getConversationReference(context.activity);
            context.adapter.createConversation(conversationRef, async context => {
              const signoutCard = CardFactory.adaptiveCard(CardTemplates.signoutCard());
              await context.sendActivity(MessageFactory.attachment(signoutCard));
            });
          } else {
            await TeamsHttp.deleteTokenFromStorage(context.activity.from.id);
            const logoutMessageCard = CardFactory.adaptiveCard(CardTemplates.logoutMessageCard());
            await context.sendActivity(MessageFactory.attachment(logoutMessageCard));
          }
          break;
        }
        case 'help':
          const helpCard = CardFactory.adaptiveCard(CardTemplates.helpCard());
          await context.sendActivity(MessageFactory.attachment(helpCard));
          break;
        default:
          if (context.activity.attachments) break; // ignore msg if with attachments
          if (context.activity.value) {
            await this.handelSubmitFromMessage(context);
            break;
          }
          const unknownCommandCard = CardFactory.adaptiveCard(CardTemplates.unknownCommandCard());
          await context.sendActivity(MessageFactory.attachment(unknownCommandCard));

      }
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const botId = context.activity.recipient.id;
      // if bot added
      if (context.activity.membersAdded.find(elem => elem.id === botId)) {
        const welcomeCard = CardFactory.adaptiveCard(CardTemplates.welcomeCard());
        await context.sendActivity(MessageFactory.attachment(welcomeCard));
      }
      await next();
    });

    // this.onMembersRemoved(async (context, next) => {
    //   console.log('on member remove');
    //   const botId = context.activity.recipient.id;
    //   const userToken = await getUserToken(context.activity.from.id);
    //   // if bot removed
    //   if (context.activity.membersRemoved.find(elem => elem.id === botId)) {
    //     const conversationId = utils.getConversationId(context.activity);
    //     const linkedProject = await TeamsHttp.getLinkedProjectFromStorage(conversationId);
    //     if (linkedProject) {
    //       await TeamsHttp.deleteLinkedProjectFromStorage(conversationId);
    //       await TeamsHttp.removeFollowerFromProject(userToken, linkedProject.oid, conversationId, context.activity.serviceUrl);
    //     }
    //   }
    //   await next();
    // });
  }

  async handleTeamsMessagingExtensionCardButtonClicked(context, cardData) {
    const actionId = cardData.actionId;
    const teamsId = context.activity.from.id;
    const userToken = await getUserToken(teamsId);

    switch (actionId) {
      case 'taskComplete_submit':
        const result = await TeamsHttp.setTaskComplete(userToken, cardData.taskOid);
        const taskCompleteCard = CardFactory.adaptiveCard(CardTemplates.taskCompleteCard(result));
        await context.sendActivity(MessageFactory.attachment(taskCompleteCard));
        break;
      case 'followTask_submit':
        const conversationId = getConversationId(context.activity);
        const serviceUrl = context.activity.serviceUrl;
        await TeamsHttp.addFollowerToTask(userToken, cardData.taskOid, conversationId, serviceUrl);
        await context.sendActivity(`Is following ${cardData.taskName} now`);
      break;
      default:
        console.log(actionId);
        await context.sendActivity('error: submit from message extension card not handled');
    }
  }

  async handelSubmitFromMessage(context) {
    await this.handleTeamsMessagingExtensionCardButtonClicked(context, context.activity.value);
  }

  async handleTeamsSigninVerifyState(context, query) {
    const verificationCode = query.state;
    const token = await utils.getUserTokenByVerificationCode(verificationCode);
    utils.addExpirationTimeForToken(token);
    if (token) {
      const teamsId = context.activity.from.id;
      await TeamsHttp.putTokenToStorage(teamsId, token);
      const loginSuccessCard = CardFactory.adaptiveCard(CardTemplates.loginSuccessCard());
      await context.sendActivity(MessageFactory.attachment(loginSuccessCard));
    } else {
      await context.sendActivity('authentication failed!!!');
    }
  }

  async handleTeamsTaskModuleFetch(context, taskModuleRequest) {
    const userToken = await getUserToken(context.activity.from.id);
    const data = taskModuleRequest.data;
    if ((!userToken || userToken.length == 0)) {
      let title, message;
      if (data.fetchId === 'addTask_fetch') {
        title = 'Add Task';
        message = 'adding a new task';
      } else if (data.fetchId === 'addComment_fetch') {
        title = 'Add Comment';
        message = 'adding a comment';
      } else if (data.fetchId === 'linkProject_fetch') {
        title = 'Link Project';
        message = 'linking a project';
      } else if (data.fetchId === 'followProject_fetch') {
        title = 'Follow Project';
        message = 'following a project';
      }
      const loginCard = CardFactory.adaptiveCard(CardTemplates.needToLoginCard(message));
      await context.sendActivity(MessageFactory.attachment(loginCard));
      return;
    }

    return await this.fetchHandler(context, data, userToken);
  }

  async fetchHandler(context, data, userToken) {
    switch (data.fetchId) {
      case 'addTask_fetch': {
        const conversationId = getConversationId(context.activity);
        const linkedProject = await TeamsHttp.getLinkedProjectFromStorage(conversationId);
        if (!linkedProject) {
          const responseCard = CardFactory.adaptiveCard(CardTemplates.needToLinkProjectButton());
          if (data.type) { // invoked by adaptive card, return a message
            await context.sendActivity(MessageFactory.attachment(responseCard));
            return;
          } else { // invoked by messaging extension, return a task module
            return createTaskInfo('Add Task', responseCard);
          }
        }

        const users = await TeamsHttp.getUsersByProjectOid(userToken, linkedProject.oid);
        const addTaskCard = CardFactory.adaptiveCard(CardTemplates.addTaskCard(linkedProject, users));
        return createTaskInfo('Add Task', addTaskCard);
      }
      case 'addComment_fetch':
        const addCommentCard = CardFactory.adaptiveCard(CardTemplates.addCommentCard(data.taskName, data.taskOid));
        return createTaskInfo('Add Comment', addCommentCard);
      case 'linkProject_fetch': {
        const conversationId = getConversationId(context.activity);
        const linkedProject = await TeamsHttp.getLinkedProjectFromStorage(conversationId);
        const allProjects = await TeamsHttp.getAllProjects(userToken);
        const linkProjectCard = CardFactory.adaptiveCard(CardTemplates.linkProjectCard(linkedProject, allProjects));
        return createTaskInfo('Link Project', linkProjectCard);
      }
      case 'followProject_fetch': {
        const allProjects = await TeamsHttp.getAllProjects(userToken);
        const followProjectCard = CardFactory.adaptiveCard(CardTemplates.followProjectCard(allProjects));
        return createTaskInfo('Follow Project', followProjectCard);
      }
      default:
        console.log(data);
        await context.sendActivity('error: fetch not handled');
    }
  }

  async handleTeamsTaskModuleSubmit(context, taskModuleRequest) {
    const teamsId = context.activity.from.id;
    const userToken = await getUserToken(teamsId);
    const data = taskModuleRequest.data;
    const actionId = data.actionId;

    switch (actionId) {
      case 'changeProject_submit':
        const originProject = data.project;
        const projects = await TeamsHttp.getAllProjects(userToken);
        const changeProjectCard = CardFactory.adaptiveCard(CardTemplates.changeProjectCard(originProject, projects));
        return createTaskInfo('Change Project', changeProjectCard);
      case 'setProject_submit':
        const selectedProject = JSON.parse(data.changeProject_input || data.originProject);
        const users = await TeamsHttp.getUsersByProjectOid(userToken, selectedProject.oid);
        const newAddProjectCard = CardFactory.adaptiveCard(CardTemplates.addTaskCard(selectedProject, users));
        return createTaskInfo('Add Task Adaptive Card', newAddProjectCard);
      case 'addTask_submit':
        const oid = data.project.oid;
        const task = {
          name: data.taskName_input,
          due: data.dueDate_input,
          description: data.description_input
        };
        if (task.name.length == 0) {
          const messageCard = CardFactory.adaptiveCard(CardTemplates.simpleMessageCard('Please input task name!'));
          return createTaskInfo('Add Task', messageCard);
        }

        if (data.assignee) {
          task.assignees = [JSON.parse(data.assignee).oid];
        }
        const respond = await TeamsHttp.addTaskToProjectByOid(userToken, task, oid);
        const taskCard = CardFactory.adaptiveCard(CardTemplates.taskCard(respond));

        await context.sendActivity('Your new task has been added to Quire.');
        await context.sendActivity(MessageFactory.attachment(taskCard));
        break;
      case 'addComment_submit':
        if (data.comment_input.length == 0) {
          const messageCard = CardFactory.adaptiveCard(CardTemplates.simpleMessageCard('Please input comment!'));
          return createTaskInfo('Add Comment', messageCard);
        }
        TeamsHttp.addCommentToTaskByOid(userToken, data.comment_input, data.taskOid)
        await context.sendActivity(`Your comment has been added to ${data.taskName}`);
        break;
      case 'linkProject_submit': {
        const teamsId = context.activity.from.id;
        const project = JSON.parse(data.linkProject_input);
        await TeamsHttp.putLinkedProjectToStorage(teamsId, project);
        await context.sendActivity(`${project.nameText} has been linked to Quire app.`);
        break;
      }
      case 'followProject_submit': {
        const conversationId = getConversationId(context.activity);
        const serviceUrl = context.activity.serviceUrl;
        if (!data.followProject_input) break;

        const project = JSON.parse(data.followProject_input);
        await TeamsHttp.addFollowerToProject(userToken, project.oid, conversationId, serviceUrl);
        await context.sendActivity(`You have successfully followed ${project.nameText}`);
        break;
      }
      case 'followTask_submit': {
        const conversationId = getConversationId(context.activity);
        const serviceUrl = context.activity.serviceUrl;
        await TeamsHttp.addFollowerToTask(userToken, data.taskOid, conversationId, serviceUrl);
        await context.sendActivity(`You have successfully followed ${data.taskName}`);
        break;
      }
      case 'unlinkProject_submit': {
        const conversationId = utils.getConversationId(context.activity);
        await TeamsHttp.deleteLinkedProjectFromStorage(conversationId);
        await context.sendActivity('This channel is unlink now');
        break;
      }
      case 'redirectToSignin_submit':
        const loginButton = CardFactory.adaptiveCard(CardTemplates.loginButton());
        return {
          composeExtension: {
            type: 'result',
            attachmentLayout: 'list',
            attachments: [ loginButton ]
          }
        };
      default:
        if (taskModuleRequest.data.fetchId === 'linkProject_fetch') {
          const conversationId = getConversationId(context.activity);
          const linkedProject = await TeamsHttp.getLinkedProjectFromStorage(conversationId);
          const allProjects = await TeamsHttp.getAllProjects(userToken);
          const linkProjectCard = CardFactory.adaptiveCard(CardTemplates.linkProjectCard(linkedProject, allProjects));
          return createTaskInfo('Link Project', linkProjectCard);
        }
        console.log(data);
        await context.sendActivity('error: submit not handled');
    }
  }

  async handleTeamsMessagingExtensionFetchTask(context, action) {
    let token;
    // handle login
    if (action.state) {
      const verificationCode = action.state;
      token = await utils.getUserTokenByVerificationCode(verificationCode);
      utils.addExpirationTimeForToken(token);
      if (token) { // if login success, put token to storage and continue work
        const teamsId = context.activity.from.id;
        await TeamsHttp.putTokenToStorage(teamsId, token);
      } else {
        return {
          composeExtension: {
            type: 'message',
            text: 'authentication failed!'
          }
        };
      }
    }

    const teamsId = context.activity.from.id;
    const userToken = token || await getUserToken(teamsId);
    if ((!userToken || userToken.length == 0)) {
      return {
        composeExtension: {
          type: 'auth',
          suggestedActions: {
            actions: [
              {
                type: 'openUrl',
                value: `https://${domainName}/bot-auth-start`,
                title: 'Sign in to this app'
              }
            ]
          }
        }
      };
    }
    const data = { fetchId: action.commandId.replace('extension', 'fetch') }
    return await this.fetchHandler(context, data, userToken);
  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    return await this.handleTeamsTaskModuleSubmit(context, action);
  }

  async handleTeamsMessagingExtensionQuery(context, query) {
    // handle login
    if (query.state) {
      const verificationCode = query.state;
      const token = await utils.getUserTokenByVerificationCode(verificationCode);
      utils.addExpirationTimeForToken(token);
      if (token) { // if login success, put token to storage and continue search
        const teamsId = context.activity.from.id;
        await TeamsHttp.putTokenToStorage(teamsId, token);
      } else {
        return {
          composeExtension: {
            type: 'message',
            text: 'authentication failed!!!'
          }
        };
      }
    }

    const userToken = await getUserToken(context.activity.from.id);
    if (!userToken)
      return {
        composeExtension: {
          type: 'auth',
          suggestedActions: {
            actions: [{
              type: 'openUrl',
              value: `https://${domainName}/bot-auth-start`,
              title: 'Log in to Quire'
            }]
          }
        }
      };

    const conversationId = utils.getConversationId(context.activity);
    const linkedProject = await TeamsHttp.getLinkedProjectFromStorage(conversationId);
    if (!linkedProject)
      return {
        composeExtension: {
          type: 'message',
          text: 'Please link a Quire project first.'
        }
      };

    const textToSearch = query.parameters[0].value;
    const attachments = [];

    let results;
    if (query.parameters[0].name === 'initialRun')
      results = await TeamsHttp.getRootTasksByOid(userToken, linkedProject.oid);
    else
      results = await TeamsHttp.searchTaskByProjectOid(userToken, textToSearch, linkedProject.oid);

    for (const task of results) {
      const adaptiveCard = CardFactory.adaptiveCard(CardTemplates.taskCardWithFollowBtn(task, linkedProject.nameText));
      adaptiveCard.preview = CardFactory.thumbnailCard(task.name, task.description);
      attachments.push(adaptiveCard);
    }

    return {
      composeExtension: {
        type: 'result',
        attachmentLayout: 'list',
        attachments: attachments
      }
    };
  }
}

function createTaskInfo(title, adaptiveCard, height, width) {
  return {
    task: {
      type: 'continue',
      value: {
        title: title,
        card: adaptiveCard,
        height: height,
        width: width
      }
    }
  };
}

module.exports.BotActivityHandler = BotActivityHandler;
