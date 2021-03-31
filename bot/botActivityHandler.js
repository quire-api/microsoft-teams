// Copyright (C) 2020 Potix Corporation. All Rights Reserved
// History: 2020/12/22 5:55 PM
// Author: charlie<charliehsieh@potix.com>

const {
  TurnContext,
  MessageFactory,
  TeamsActivityHandler,
  CardFactory,
} = require('botbuilder');
const { CardTemplates } = require('../model/cardtemplates');
const { QuireApi } = require('../utils/quireApi');
const dbAccess = require('../db/dbAccess');
const utils = require('../utils/utils');
const { QuireMessages } = require('../utils/quireMessages');
const domainName = process.env.DomainName;

class BotActivityHandler extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      TurnContext.removeRecipientMention(context.activity);
      const command = new String(context.activity.text).trim().toLocaleLowerCase();
      const isLogin = await utils.isUserLogin(context.activity.from.id);
      if (context.activity.value) {
        return await this.handleTeamsTaskModuleSubmit(context, {data: context.activity.value});
      }

      switch (command) {
        case 'help':
          await context.sendActivity(MessageFactory.attachment(
                CardTemplates.helpCard()));
          break;
        case 'take a tour':
          await context.sendActivity(MessageFactory.carousel(
            CardTemplates.tourCard()));
          break;
        case 'login': {
          const conversationType = context.activity.conversation.conversationType;
          const conversationRef = TurnContext.getConversationReference(context.activity);
          if (conversationType === 'groupChat' || conversationType === 'channel') {
            await context.sendActivity(`Thanks ${context.activity.from.name}, I've sent you a direct message to help you do this. If you don't see the message, try adding the Quire app first`);
          }

          let returnMessage;
          if (isLogin) {
            returnMessage = MessageFactory.text('Hey, you’re already logged in.');
          } else {
            const loginButton = CardTemplates.loginButton();
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
            await context.sendActivity(`Thanks ${context.activity.from.name}, I've sent you a direct message to help you do this. If you don't see the message, try adding the Quire app first`);
            const conversationRef = TurnContext.getConversationReference(context.activity);
            context.adapter.createConversation(conversationRef, async context => {
              const signoutCard = CardTemplates.signoutCard();
              await context.sendActivity(MessageFactory.attachment(signoutCard));
            });
          } else {
            if (isLogin) {
              await dbAccess.deleteToken(context.activity.from.id);
              const logoutMessageCard = CardTemplates.logoutMessageCard();
              await context.sendActivity(MessageFactory.attachment(logoutMessageCard));
            } else {
              await context.sendActivity(MessageFactory.text('You are already logged out. You can always login again. See you later!'));
            }
          }
          break;
        }
        default:
          if (isLogin) {
            await this.handleTeamsCommands(context, command);
          } else {
            const desc = QuireMessages.getCommandDescriptions(command);
            await context.sendActivity(MessageFactory.attachment(
                CardTemplates.needToLoginCard(desc)));
          }
      }
      
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const conversationType = context.activity.conversation.conversationType;
      // if bot added to a personal chat, send proactive welcome message
      if (conversationType === 'personal') {
        const welcomeCard = CardTemplates.welcomeCard();
        await context.sendActivity(MessageFactory.attachment(welcomeCard));
      } else if (conversationType === 'groupChat' || conversationType === 'channel') {
        const botId = context.activity.recipient.id;
        if (context.activity.membersAdded.find(elem => elem.id === botId)) {
          dbAccess.addToTeamList(utils.getConversationId(context.activity));
        }
      }
      await next();
    });

    this.onMembersRemoved(async (context, next) => {
      const conversationType = context.activity.conversation.conversationType;
      if (conversationType === 'groupChat' || conversationType === 'channel') {
        const botId = context.activity.recipient.id;
        if (context.activity.membersRemoved.find(elem => elem.id === botId)) {
          dbAccess.removeFromTeamList(utils.getConversationId(context.activity));
        }
      }
      await next();
    });
  }

  async handleTeamsCommands(context, command) {
    switch (command) {
      case 'add task':
      case 'create task':
        let respondCard;
        await context.sendActivity(MessageFactory.attachment(
            CardTemplates.addTaskButton()));
        break;
      case 'link project':
        await context.sendActivity(MessageFactory.attachment(
            CardTemplates.linkProjectButton()));
        break;
      case 'unlink project':
        const conversationId = utils.getConversationId(context.activity);
        const linkedProject = await dbAccess.getLinkedProject(conversationId);
        if (!linkedProject) {
          await context.sendActivity(`You haven't linked any project to this channel yet.`);
          break;
        }

        dbAccess.removeLinkedProject(conversationId);
        await context.sendActivity(`You have unlinked ${linkedProject.nameText} from this channel.`);
        break;
      case 'follow project':
        await context.sendActivity(MessageFactory.attachment(
            CardTemplates.followProjectButton()));
        break;
      case 'unfollow project':
        await context.sendActivity(MessageFactory.attachment(
            CardTemplates.unfollowProjectButton()));
        break;
      default:
        if (context.activity.attachments) break; // ignore msg if with attachments
        await context.sendActivity(MessageFactory.attachment(
          CardTemplates.unknownCommandCard()));
    }
  }

  async handleTeamsMessagingExtensionCardButtonClicked(context, cardData) {
    const conversationType = context.activity.conversation.conversationType;
    if (!conversationType) return;

    const actionId = cardData.actionId;
    const teamsId = context.activity.from.id;
    const isLogin = await utils.isUserLogin(teamsId);
    if (!isLogin) {
        const desc = QuireMessages.getButtonLabel(actionId);
        await context.sendActivity(MessageFactory.attachment(
            CardTemplates.needToLoginCard(desc)));
        return;
    }

    const userToken = await dbAccess.getToken(teamsId);

    try {
      switch (actionId) {
        case 'taskComplete_submit':
          const result = await QuireApi.setTaskComplete(userToken, cardData.taskOid);
          const taskCompleteCard = CardTemplates.taskCompleteCard(result);
          await context.sendActivity(MessageFactory.attachment(taskCompleteCard));
          break;
        case 'followTask_submit': {
          const conversationId = utils.getConversationId(context.activity);
          const serviceUrl = context.activity.serviceUrl;
          await QuireApi.addFollowerToTask(userToken, cardData.taskOid, conversationId, serviceUrl);
          await context.sendActivity(`You have got this channel to follow ${cardData.taskName}.`);
          break;
        }
        default:
          console.log(actionId);
          await context.sendActivity('error: submit from message extension card not handled');
      }
    } catch (error) {
      if (error.hasNoPermission) {
        await context.sendActivity('You do not have permission to perform this action. Please contact your Admin.');
      } else if (error.connectionError) {
        await context.sendActivity('Sorry, Quire is temporary unavailable. We will be back shortly.');
      } else if (error.notFound) {
        await context.sendActivity(`Task ${cardData.nameText} not found. Please check again in Quire.`);
      } else {
        throw error;
      }
    }
  }

  async handleTeamsSigninVerifyState(context, query) {
    const verificationCode = query.state;
    const token = await utils.getUserTokenByVerificationCode(verificationCode);
    utils.addExpirationTimeForToken(token);
    if (token) {
      const teamsId = context.activity.from.id;
      dbAccess.putToken(teamsId, token);
      const loginSuccessCard = CardTemplates.loginSuccessCard();
      await context.sendActivity(MessageFactory.attachment(loginSuccessCard));
    } else {
      await context.sendActivity('Authentication failed!!!');
    }
  }

  async handleTeamsTaskModuleFetch(context, taskModuleRequest) {
    const data = taskModuleRequest.data;
    const teamsId = context.activity.from.id;
    let userToken;
    try {
      userToken = data.token || await dbAccess.getToken(teamsId);
      if (userToken)
        return await this.fetchHandler(context, data, userToken);
      else
        return await this.sendPleaseLoginCard(context, data);
    } catch (error) {
      if (error.response && error.response.status === 401) {
        // try to refresh token and fetch again
        if (userToken && !data.token) {
          const token = await QuireApi.refreshAndStoreToken(teamsId, userToken);
          if (!token.isInvalidToken) {
            data.token = token;
            return await this.handleTeamsTaskModuleFetch(context, taskModuleRequest);
          }
        }

        // refresh token failed, send 'Please login' message
        this.sendPleaseLoginCard(context, data);
      } else {
        let message;
        if (error.hasNoPermission) {
          message = 'You do not have permission to perform this action. Please contact your Admin.';
        } else if (error.connectionError) {
          message = 'Sorry, Quire is temporary unavailable. We will be back shortly.';
        } else if (error.notFound) {
          message = `Task ${taskModuleRequest.data.taskName} not found. Please check again in Quire.`;
        }
        return createTaskInfo('Error', CardTemplates.simpleMessageCard(message));
      }
    }
  }

  async sendPleaseLoginCard(context, data) {
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
    } else if (data.fetchId === 'unfollowProject_fetch') {
      title = 'Unfollow Project';
      message = 'unfollow project';
    } else if (data.fetchId === 'followTask_submit') {
      title = 'Follow Task';
      message = 'following a task';
    } else if (data.fetchId === 'taskComplete_submit') {
      title = 'Complete Task';
      message = 'completing a task';
    }
    if (context.activity.conversation.conversationType === 'personal') {
      const loginCard = CardTemplates.needToLoginCard(message);
      await context.sendActivity(MessageFactory.attachment(loginCard));
    } else {
      message = `You need to log into your Quire account before ${message}`
      const card = CardTemplates.simpleMessageCard(message);
      return createTaskInfo(title, card);
    }
  }

  async fetchHandler(context, data, userToken) {
    const conversationType = context.activity.conversation.conversationType;    
    if ((conversationType === 'groupChat' || conversationType === 'channel') && !await dbAccess.isChannelMember(utils.getConversationId(context.activity))) {
      return createTaskInfo('Error', CardTemplates.pleaseAddBotToChannelCard(conversationType));
    }

    switch (data.fetchId) {
      case 'addTask_fetch': {
        const conversationId = utils.getConversationId(context.activity);
        const linkedProject = await dbAccess.getLinkedProject(conversationId);
        if (!linkedProject) {
          const responseCard = CardTemplates.needToLinkProjectButton();
          if (data.type) { // invoked by adaptive card, return a message
            await context.sendActivity(MessageFactory.attachment(responseCard));
            return;
          } else { // invoked by messaging extension, return a task module
            return createTaskInfo('Add Task', responseCard);
          }
        }

        const users = await QuireApi.getUsersByProjectOid(userToken, linkedProject.oid);
        const addTaskCard = CardTemplates.addTaskCard(linkedProject, users);
        return createTaskInfo('Add Task', addTaskCard);
      }
      case 'addComment_fetch':
        const addCommentCard = CardTemplates.addCommentCard(data.taskName, data.taskOid);
        return createTaskInfo('Add Comment', addCommentCard);
      case 'linkProject_fetch': {
        const conversationId = utils.getConversationId(context.activity);
        const linkedProject = await dbAccess.getLinkedProject(conversationId);
        const allProjects = await QuireApi.getAllProjects(userToken);
        const linkProjectCard = CardTemplates.linkProjectCard(linkedProject, allProjects);
        return createTaskInfo('Link Project', linkProjectCard);
      }
      case 'unlinkProject_fetch': {
        const conversationId = utils.getConversationId(context.activity);
        const linkedProject = await dbAccess.getLinkedProject(conversationId);
        let message;
        if (linkedProject) {
          await dbAccess.removeLinkedProject(conversationId);
          message = `You have unlinked ${linkedProject.nameText} from this channel.`;
        } else {
          message = "You haven't linked any project to this channel yet.";
        }

        return createTaskInfo('Unlink Project', CardTemplates.simpleMessageCard(message));
      }
      case 'followProject_fetch': {
        const allProjects = await QuireApi.getAllProjects(userToken);
        allProjects.shift(); // remove my task
        const followProjectCard = CardTemplates.followProjectCard(allProjects);
        return createTaskInfo('Follow Project', followProjectCard);
      }
      case 'unfollowProject_fetch': {
        const isInvokedByMessageExtension = data.type !== 'task/fetch';
        const conversationId = utils.getConversationId(context.activity);
        const followedProjectList = await dbAccess.getFollowedProjectList(conversationId);
        if (followedProjectList.length == 0) {
          const message = "You haven't got this channel to follow any project yet."
          if (isInvokedByMessageExtension) {
            return createTaskInfo('Unfollow Project', CardTemplates.simpleMessageCard(message));
          } else {
            await context.sendActivity(message);
          }
          break;
        }
        return createTaskInfo('Unfollow Project', CardTemplates.unfollowProjectCard(followedProjectList));
      }
      case 'taskComplete_submit': {
        const task = await QuireApi.getTaskByOid(userToken, data.taskOid);
        var message;
        if (!task) {
          message = 'Task not found.';
        } else {
          const result = await QuireApi.setTaskComplete(userToken, data.taskOid);
          message = `${result.nameText} has been completed.`
        }
        await context.sendActivity(message);
        break;
      }
      case 'followTask_submit': {
        const conversationId = utils.getConversationId(context.activity);
        const serviceUrl = context.activity.serviceUrl;
        const respond = await QuireApi.addFollowerToTask(userToken, data.taskOid, conversationId, serviceUrl);
        if (respond.hasNoPermission) {
          const messageCard = CardTemplates.simpleMessageCard('You do not have permission to perform this action. Please contact your Admin.');
          return createTaskInfo('Follow Task', messageCard);
        }
        const messageCard = 
            CardTemplates.simpleMessageCard(`You have successfully followed ${data.taskName}`);
        return createTaskInfo('Follow Task', messageCard);
      }
      default:
        console.log(data);
        await context.sendActivity('error: fetch not handled');
    }
  }

  async handleTeamsTaskModuleSubmit(context, taskModuleRequest) {
    const data = taskModuleRequest.data;
    const teamsId = context.activity.from.id;
    let userToken;
    try {
      userToken = data.token || await dbAccess.getToken(teamsId);
      return await this.handleSubmit(context, data, userToken);
    } catch (error) {
      if (error.code === 'BotNotInConversationRoster') {
        return; // ignore
      } else if (error.response && error.response.status === 401) {
        const token = await QuireApi.refreshAndStoreToken(teamsId, userToken);
        if (token.isInvalidToken) {
          let message;
          switch (data.actionId) {
            case 'changeProject_submit':
              message = 'changing project';
              break;
            case 'setProject_submit':
              message = 'setting project';
              break;
            case 'addTask_submit':
              message = 'adding a task';
              break;
            case 'addComment_submit':
              message = 'adding a comment';
              break;
            case 'taskComplete_submit':
              message = 'completing a task';
              break;
            case 'linkProject_submit':
              message = 'linking a project';
              break;
            case 'followProject_submit':
              message = 'following a project';
              break;
            case 'followTask_submit':
              message = 'following a task';
              break;
            case 'unfollowProject_submit':
              message = 'unfollowing a project';
              break;
          }
          const needToLoginCard = CardTemplates.needToLoginCard(message);
          await context.sendActivity(MessageFactory.attachment(needToLoginCard));
          return;
        }

        taskModuleRequest.token = token;
        return await this.handleTeamsTaskModuleSubmit(context, taskModuleRequest);
      } else {
        let message;
        if (error.hasNoPermission) {
          message = 'You do not have permission to perform this action. Please contact your Admin.';
        } else if (error.connectionError) {
          message = 'Sorry, Quire is temporary unavailable. We will be back shortly.';
        } else if (error.notFound) {
          message = `Task ${taskModuleRequest.data.taskName} not found. Please check again in Quire.`;
        }

        return createTaskInfo('Error', CardTemplates.simpleMessageCard(message));
      }
    }
  }

  async handleSubmit(context, data, userToken) {
    const actionId = data.actionId;

    switch (actionId) {
      case 'changeProject_submit':
        const originProject = data.project;
        const projects = await QuireApi.getAllProjects(userToken);
        const changeProjectCard = CardTemplates.changeProjectCard(originProject, projects);
        return createTaskInfo('Change Project', changeProjectCard);
      case 'setProject_submit':
        const selectedProject = JSON.parse(data.changeProject_input || data.originProject);
        const users = await QuireApi.getUsersByProjectOid(userToken, selectedProject.oid);
        const newAddProjectCard = CardTemplates.addTaskCard(selectedProject, users);
        return createTaskInfo('Add Task', newAddProjectCard);
      case 'addTask_submit':
        const oid = data.project.oid;
        const task = {
          name: data.taskName_input,
          due: data.dueDate_input,
          description: data.description_input
        };
        if (task.name.length == 0) {
          const messageCard = CardTemplates.simpleMessageCard('Task name cannot be empty.');
          return createTaskInfo('Add Task', messageCard);
        }

        if (data.assignee) {
          task.assignees = [JSON.parse(data.assignee).oid];
        }
        const conversationType = context.activity.conversation.conversationType;
        const respond = await QuireApi.addTaskToProjectByOid(userToken, task, oid);
        if (respond.hasNoPermission) {
          const messageCard = CardTemplates.simpleMessageCard('You do not have permission to perform this action. Please contact your Admin.');
          return createTaskInfo('Add Task', messageCard);
        }

        const taskCard = 
            CardTemplates.taskCard(respond, data.project.nameText, conversationType);
        await this.sendMessageToMember(context, async (t) => {
          await t.sendActivity(`Your new task ${data.taskName_input} has been added to Quire.`);
          await t.sendActivity(MessageFactory.attachment(taskCard));
        });

        if (conversationType !== 'personal')
          await context.sendActivity(MessageFactory.attachment(taskCard));

        break;
      case 'addComment_submit': {
        if (data.comment_input.length == 0) {
          const messageCard = CardTemplates.simpleMessageCard('Task comment cannot be empty.');
          return createTaskInfo('Add Comment', messageCard);
        }
        const task = await QuireApi.addCommentToTaskByOid(userToken, data.comment_input, data.taskOid);
        if (task.hasNoPermission) {
          const messageCard = CardTemplates.simpleMessageCard('You do not have permission to perform this action. Please contact your Admin.');
          return createTaskInfo('Follow Project', messageCard);
        }
        const commentCard = CardTemplates.commentCard(context.activity.from.name, task.owner.name, task.description, task.url);
        
        await context.sendActivity(MessageFactory.attachment(commentCard));
        break;
      }
      case 'taskComplete_submit': {
        const task = await QuireApi.getTaskByOid(userToken, data.taskOid);
        var message;
        if (!task) {
          message = 'Task not found.';
        } else {
          const result = await QuireApi.setTaskComplete(userToken, data.taskOid);
          message = `${result.nameText} has been completed.`
        }
        await context.sendActivity(message);
        break;
      }
      case 'linkProject_submit': {
        if (!data.linkProject_input)
          return createTaskInfo('Link Project', CardTemplates.simpleMessageCard('Project name cannot be empty'));

        const id = utils.getConversationId(context.activity);
        const project = JSON.parse(data.linkProject_input);
        dbAccess.putLinkedProject(id, project);
        const message = `You have successfully linked ${project.nameText} to this channel.`;
        if (context.activity.conversation.conversationType === 'personal') {
          await context.sendActivity(message);
          break;
        }
        const messageCard = CardTemplates.simpleMessageCard(message);
        return createTaskInfo('Link Project', messageCard);
      }
      case 'followProject_submit': {
        const conversationId = utils.getConversationId(context.activity);
        const serviceUrl = context.activity.serviceUrl;
        if (!data.followProject_input)
          return createTaskInfo('Follow Project', CardTemplates.simpleMessageCard('Project name cannot be empty'));

        const project = JSON.parse(data.followProject_input);
        const respond = await QuireApi.addFollowerToProject(userToken, project.oid, conversationId, serviceUrl);
        if (respond.hasNoPermission) {
          const messageCard = CardTemplates.simpleMessageCard('You do not have permission to perform this action. Please contact your Admin.');
          return createTaskInfo('Follow Project', messageCard);
        }

        dbAccess.addToFollowedProjectList(project.oid, conversationId, project.nameText);
        const message = `You have got this channel to follow ${project.nameText}.`;
        if (context.activity.conversation.conversationType === 'personal') {
          await context.sendActivity(message);
          break;
        }
        const messageCard = CardTemplates.simpleMessageCard(message);
        return createTaskInfo('Follow Project', messageCard);
      }
      case 'followTask_submit': {
        const conversationId = utils.getConversationId(context.activity);
        const serviceUrl = context.activity.serviceUrl;
        const respond = await QuireApi.addFollowerToTask(userToken, data.taskOid, conversationId, serviceUrl);
        if (respond.hasNoPermission) {
          const messageCard = CardTemplates.simpleMessageCard('You do not have permission to perform this action. Please contact your Admin.');
          return createTaskInfo('Follow Task', messageCard);
        }
        await context.sendActivity(`You have successfully followed ${data.taskName}`);
        break;
      }
      case 'unfollowProject_submit': {
        if (!data.unfollowProject_input)
          return createTaskInfo('Unfollow Project', CardTemplates.simpleMessageCard('Project name cannot be empty'));

        const project = JSON.parse(data.unfollowProject_input);
        const conversationId = utils.getConversationId(context.activity);
        const serviceUrl = context.activity.serviceUrl;
        const respond = await QuireApi.removeFollowerFromProject(userToken, project.oid, conversationId, serviceUrl);
        if (respond.hasNoPermission) {
          const messageCard = CardTemplates.simpleMessageCard('You do not have permission to perform this action. Please contact your Admin.');
          return createTaskInfo('Unfollow Task', messageCard);
        }
        dbAccess.removeFromFollowedProjectList(project.oid, conversationId);
        await context.sendActivity(`You have got this channel to unfollow ${project.nameText}.`);
        break;
      }
      case 'redirectToSignin_submit':
        const loginButton = CardTemplates.loginButton();
        return {
          composeExtension: {
            type: 'result',
            attachmentLayout: 'list',
            attachments: [ loginButton ]
          }
        };
      default:
        if (data.fetchId === 'linkProject_fetch') {
          const conversationId = utils.getConversationId(context.activity);
          const linkedProject = await dbAccess.getLinkedProject(conversationId);
          const allProjects = await QuireApi.getAllProjects(userToken);
          const linkProjectCard = CardTemplates.linkProjectCard(linkedProject, allProjects);
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
      if (token) { // if login success, put token to storage and continue work
        const teamsId = context.activity.from.id;
        dbAccess.putToken(teamsId, token);
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
    const userToken = token || action.token || await dbAccess.getToken(teamsId);
    const loginAction = {
      composeExtension: {
        type: 'auth',
        suggestedActions: {
          actions: [
            {
              type: 'openUrl',
              value: `${domainName}/bot-auth-start`,
              title: 'Sign in to this app'
            }
          ]
        }
      }
    };
    if (!userToken)
      return loginAction;
    
    const data = { fetchId: action.commandId.replace('extension', 'fetch') }
    try {
      return await this.fetchHandler(context, data, userToken);
    } catch (error) {
      if (error.isAxiosError) {
        if (error.response.status === 401) {
          const token = await QuireApi.refreshAndStoreToken(teamsId, userToken);
          if (token.isInvalidToken)
            return loginAction;
  
          action.token = token;
          return await this.handleTeamsMessagingExtensionFetchTask(context, action);
        } else {
          const conversationId = utils.getConversationId(context.activity);
          const linkedProject = await dbAccess.getLinkedProject(conversationId);
          if (error.response.status == 403) {  
            const messageCard = CardTemplates.simpleMessageCard(
              `Sorry, you do not have permission to access project ${linkedProject.nameText}. Please contact Quire project admin.`);
            return createTaskInfo('Add Task', messageCard);
          } else if (error.response.status == 404) {
            const messageCard = CardTemplates.simpleMessageCard(
              `Project ${linkedProject.nameText}: not found.`);
            return createTaskInfo('Add Task', messageCard);
          }
        }
      }
      throw error;
    }
  }

  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    const teamsId = context.activity.from.id;
    let token;
    if (action.state) {
      const verificationCode = action.state;
      token = await utils.getUserTokenByVerificationCode(verificationCode);
      utils.addExpirationTimeForToken(token);
      if (token) { // if login success, put token to storage and continue search
        dbAccess.putToken(teamsId, token);
      } else {
        return {
          composeExtension: {
            type: 'message',
            text: 'authentication failed!!!'
          }
        };
      }
    }

    const userToken = token || action.token || await dbAccess.getToken(teamsId);
    const data = action.data;
    const loginAction = {
      composeExtension: {
        type: 'auth',
        suggestedActions: {
          actions: [{
            type: 'openUrl',
            value: `${domainName}/bot-auth-start`,
            title: 'Log in to Quire'
          }]
        }
      }
    };
    if (!userToken)
      return loginAction;

    try {
      return await this.handleSubmit(context, data, userToken);
    } catch (error) {
      if (!(error.isAxiosError && error.response.status === 401))
        throw error;

      const token = await QuireApi.refreshAndStoreToken(teamsId, userToken);
      if (token.isInvalidToken)
        return loginAction;

      action.token = token;
      return await this.handleTeamsMessagingExtensionSubmitAction(context, action);
    }
  }

  async handleTeamsMessagingExtensionQuery(context, query) {
    const teamsId = context.activity.from.id;
    let token;
    // handle login
    if (query.state) {
      const verificationCode = query.state;
      token = await utils.getUserTokenByVerificationCode(verificationCode);
      utils.addExpirationTimeForToken(token);
      if (token) { // if login success, put token to storage and continue search
        dbAccess.putToken(teamsId, token);
      } else {
        return {
          composeExtension: {
            type: 'message',
            text: 'authentication failed!!!'
          }
        };
      }
    }

    const userToken = token || query.token || await dbAccess.getToken(teamsId);
    const loginAction = {
      composeExtension: {
        type: 'auth',
        suggestedActions: {
          actions: [{
            type: 'openUrl',
            value: `${domainName}/bot-auth-start`,
            title: 'Log in to Quire'
          }]
        }
      }
    };
    if (!userToken)
      return loginAction;

    const conversationId = utils.getConversationId(context.activity);
    const linkedProject = await dbAccess.getLinkedProject(conversationId);
    try {
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
        results = await QuireApi.getRootTasksByOid(userToken, linkedProject.oid);
      else
        results = await QuireApi.searchTaskByProjectOid(userToken, textToSearch, linkedProject.oid);

      const conversationType = context.activity.conversation.conversationType;
      for (const task of results) {
        if (task.status.value == 100) continue;
        const adaptiveCard = 
            CardTemplates.taskCardWithFollowBtn(task, linkedProject.nameText, conversationType);
        adaptiveCard.preview = CardFactory.thumbnailCard(task.nameText, task.description);
        attachments.push(adaptiveCard);
      }

      return {
        composeExtension: {
          type: 'result',
          attachmentLayout: 'list',
          attachments: attachments
        }
      };
    } catch (error) {
      if (error.timeout)
        return {
          composeExtension: {
            type: 'message',
            text: 'Sorry, your search session is timeout. Please try again.'
          }
        };

      if (error.isAxiosError) {
        // 401 Invalid or expired token, refresh token and try again
        if (error.response.status === 401) {
          const token = await QuireApi.refreshAndStoreToken(teamsId, userToken);
          if (token.isInvalidToken)
            return loginAction;

          query.token = token;
          return await this.handleTeamsMessagingExtensionQuery(context, query);

        // 403 Not authorized to access the resource.
        } else if (error.response.status == 403) {
          return {
            composeExtension: {
              type: 'message',
              text: `Sorry, you do not have permission to access project ${linkedProject.nameText}. Please contact Quire project admin.`
            }
          };
        // 404 The specified resource could not be found.
        } else if (error.response.status == 404) {
          return {
            composeExtension: {
              type: 'message',
              text: `Project ${linkedProject.nameText}: not found.`
            }
          };
        }
      }
      throw error;
    }
  }

  async sendMessageToMember(context, callback) {
    let ref = TurnContext.getConversationReference(context.activity);
    ref.user = {
      id: context.activity.from.id,
      aadObjectId: context.activity.from.aadObjectId,
      tenantId: context.activity.conversation.tenantId
    }

    await context.adapter.createConversation(ref, async (t1) => {
      const ref2 = TurnContext.getConversationReference(t1.activity);
      await t1.adapter.continueConversation(ref2, async (t2) => {
        await callback(t2);
      });
    });
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
