// Copyright (C) 2020 Potix Corporation. All Rights Reserved
// History: 2020/12/22 5:55 PM
// Author: charlie<charliehsieh@potix.com>

const { CardFactory } = require("botbuilder");
const utils = require("../utils/utils");
const taskDescriptionLimit = 4000;

class CardTemplates {
  static welcomeCard() {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: 'Hi there👋  I\'m Quire Bot and I\'m here at your service!\n\nI am here to make your work much easier and faster! You can get real-time updates from your projects in Quire and manage your task list right from here at Teams!\n\n**Here are a couple things that I can do:**\n\n- Add new tasks to Quire projects.\n- Assign tasks to team members, set dates and more.\n\nTo get you started, let\'s log in your Quire workspace.',
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Log in to Quire',
          data: {
            msteams: {
              type: 'signin',
              value: `${process.env.DomainName}/bot-auth-start`
            }
          }
        },
        {
          type: 'Action.OpenUrl',
          title: 'Sign up',
          url: 'https://quire.io/signup?continue=https://quire.io/r/integra/microsoft-teams/signup/confirm'
        }
      ]
    });
  }

  static loginSuccessCard() {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.3',
      body: [
        {
          type: 'TextBlock',
          text: 'Great, you’re logged in 🎉\n\nFirst things first, you need to decide which project you would you like to link your Microsoft Teams.',
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Link a project',
          data: {
            fetchId: 'linkProject_fetch',
            msteams: { type: 'task/fetch' }
          }
        },
        {
          type: 'Action.Submit',
          title: 'Take a tour',
          data: {
            msteams: {
              type: 'imBack',
              value: 'Take a tour'
            }
          }
        }
      ]
    });
  }

  static needToLoginCard(text) {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: `You need to log into your Quire account before ${text}.`,
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Log in to Quire',
          data: {
            msteams: {
              type: 'signin',
              value: `${process.env.DomainName}/bot-auth-start`
            }
          }
        }
      ]
    });
  }

  static addTaskButton() {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: 'Click the below button to add a new Quire task.',
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Add task',
          data: {
            fetchId: 'addTask_fetch',
            msteams: {
              type: 'task/fetch'
            }
          }
        }
      ]
    });
  }

  static addTaskCard(project, users) {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column',
              width: 'auto',
              verticalContentAlignment: 'Center',
              items: [
                {
                  type: 'FactSet',
                  facts: [
                    {
                      title: 'Project',
                      value: project.nameText
                    }
                  ]
                }
              ]
            },
            {
              type: 'Column',
              width: 'stretch',
              items: [
                {
                  type: 'ActionSet',
                  actions: [
                    {
                      type: 'Action.Submit',
                      title: 'Change project',
                      data: { actionId: 'changeProject_submit', project: project }
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          type: 'TextBlock',
          text: 'Task name',
          wrap: true
        },
        {
          type: 'Input.Text',
          id: 'taskName_input',
          placeholder: 'Task name'
        },
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column',
              items: [
                {
                  type: 'TextBlock',
                  text: 'Assignee',
                  wrap: true
                },
                {
                  type: 'Input.ChoiceSet',
                  id: 'assignee',
                  placeholder: 'Select assignee',
                  choices: utils.itemsToChoices(users)
                }
              ]
            },
            {
              type: 'Column',
              items: [
                {
                  type: 'TextBlock',
                  text: 'Due date',
                  wrap: true
                },
                {
                  type: 'Input.Date',
                  id: 'dueDate_input'
                }
              ]
            }
          ]
        },
        {
          type: 'TextBlock',
          text: 'Description',
          wrap: true
        },
        {
          type: 'Input.Text',
          id: 'description_input',
          placeholder: 'Task description',
          IsMultiline: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Add task',
          data: { actionId: 'addTask_submit', project: project }
        }
      ]
    });
  }

  static changeProjectCard(originProject, projects) {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: 'Select a project to add task',
          wrap: true
        },
        {
          type: 'Input.ChoiceSet',
          id: 'changeProject_input',
          value: JSON.stringify({ oid:originProject.oid, nameText: originProject.nameText }),
          choices: utils.projectsToChoices(projects)
        }
      ],
      actions: [{
        type: 'Action.Submit',
        title: 'OK',
        data: { actionId: 'setProject_submit', originProject: originProject }
      }]
    });
  }

  static addHeaderToCard(card, headerMessage) {
    card.body.unshift({
      type: 'Container',
      bleed: true,
      style: 'emphasis',
      items: [
        {
          type: 'TextBlock',
          text: headerMessage,
          wrap: true
        }
      ]
    });
    return card;
  }

  static _taskCard(task, projectName, conversationType) {
    let descriptionText = task.descriptionText;
    if (descriptionText.length > taskDescriptionLimit)
      descriptionText = descriptionText.substr(0, taskDescriptionLimit) + '...';

    if (task.nameText.length == 0)
      task.nameText = '(empty)';

    return {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          size: 'large',
          text: task.nameText,
          wrap: true
        },
        {
          type: 'FactSet',
          facts: [
            {
              title: 'Assigned to',
              value: (task.assignees[0] || {}).name || 'None'
            },
            {
              title: 'Due date',
              value: task.due || 'Not set'
            },
            {
              title: 'Project',
              value: projectName
            }
          ]
        },
        {
          type: 'TextBlock',
          text: descriptionText,
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: 'View in Quire',
          url: task.url
        },
        {
          type: 'Action.Submit',
          title: 'Add comment',
          data: {
            fetchId: 'addComment_fetch',
            taskOid: task.oid,
            taskName: task.nameText,
            msteams: {
              type: 'task/fetch'
            }
          }
        },
       /*{
         type: 'Action.Submit',
         title: 'Complete task',
         data: {
           actionId: 'taskComplete_submit',
           fetchId: 'taskComplete_submit',
           taskOid: task.oid,
           taskName: task.nameText,
           msteams: conversationType === 'personal' ?
               null : { type: 'task/fetch' }
         }
       }*/
      ]
    };
  }

  static taskCard(task, projectName, conversationType) {
    return CardFactory.adaptiveCard(this._taskCard(task, projectName, conversationType));
  }

  static taskCardWithFollowBtn(task, projectName, conversationType) {
    const taskCard = this._taskCard(task, projectName, conversationType);
    taskCard.actions.push({
      type: 'Action.Submit',
          title: 'Follow task',
          data: {
            actionId: 'followTask_submit',
            fetchId: 'followTask_submit',
            taskOid: task.oid,
            taskName: task.nameText,
            msteams: conversationType === 'personal' ?
                null : { type: 'task/fetch' }
          }
    });
    return CardFactory.adaptiveCard(taskCard);
  }

  static addCommentCard(taskName, taskOid) {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: `Add a comment to ${taskName}`,
          wrap: true
        },
        {
          type: 'Input.Text',
          placeholder: 'Write some comment here...',
          id: 'comment_input',
          isMultiline: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Add comment',
          data: { actionId: 'addComment_submit', taskOid: taskOid, taskName: taskName }
        }
      ]
    });
  }

  static commentCard(userName, taskName, comment, taskUrl) {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: `${userName} commented on ${taskName}`,
          wrap: true
        },
        {
          type: 'TextBlock',
          text: `"${comment}"`,
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: 'View in Quire',
          url: taskUrl
        }
      ]
    });
  }

  static _linkProjectCard(message) {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: message,
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Link project',
          data: {
            fetchId: 'linkProject_fetch',
            msteams: {
              type: 'task/fetch'
            }
          }
        }
      ]
    });
  }

  static linkProjectButton() {
    return this._linkProjectCard('Click the below button to link a Quire project to this channel.');
  }

  static needToLinkProjectButton() {
    return this._linkProjectCard('Sorry, you need to link a project in Quire before adding a new task.');
  }

  static linkProjectCard(originProject, projects) {
    let project;
    if (originProject)
      project = originProject;
    else
      project = {};


    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: 'Please select a Quire project to link to this channel.',
          wrap: true
        },
        {
          type: 'Input.ChoiceSet',
          id: 'linkProject_input',
          value: JSON.stringify({ oid: project.oid, nameText: project.nameText }),
          choices: utils.projectsToChoices(projects)
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Link project',
          data: { actionId: 'linkProject_submit' }
        }
      ]
    });
  }

  static followProjectButton() {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: 'Click the below button to follow a Quire project.',
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Follow project',
          data: {
            fetchId: 'followProject_fetch',
            msteams: {
              type: 'task/fetch'
            }
          }
        }
      ]
    });
  }

  static followProjectCard(projects) {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: 'Please select a project to follow.',
          wrap: true
        },
        {
          type: 'Input.ChoiceSet',
          id: 'followProject_input',
          choices: utils.projectsToChoices(projects),
          placeholder: 'Select a project'
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Follow project',
          data: { actionId: 'followProject_submit' }
        }
      ]
    });
  }

  static unfollowProjectButton() {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: 'Click the below button to unfollow a Quire project.',
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Unfollow project',
          data: {
            fetchId: 'unfollowProject_fetch',
            msteams: {
              type: 'task/fetch'
            }
          }
        }
      ]
    });
  }

  static unfollowProjectCard(followedProjectList) {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: 'Please select a project to unfollow',
          wrap: true
        },
        {
          type: 'Input.ChoiceSet',
          id: 'unfollowProject_input',
          choices: utils.followedProjectListToChoices(followedProjectList),
          placeholder: 'Select a project'
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Unfollow project',
          data: { actionId: 'unfollowProject_submit' }
        }
      ]
    });
  }

  static signoutCard() {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: 'Press to confirm logout',
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Logout',
          data: {
            msteams: {
              type: 'imBack',
              value: 'Logout'
            }
          }
        }
      ]
    });
  }

  static _loginCard(message) {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: message,
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Log in to Quire',
          data: {
            msteams: {
              type: 'signin',
              value: `${process.env.DomainName}/bot-auth-start`
            }
          }
        }
      ]
    });
  }

  static loginButton() {
    return this._loginCard('Click the button to log in to your Quire account.')
  }

  static needToLoginButton(message) {
    return this._loginCard(`You need to log into your Quire account before ${message}`);
  }

  static logoutMessageCard() {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: 'We\'ve logged you out.\n\nYou can always login again later. Bye for now! 👋',
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Log in to Quire',
          data: {
            msteams: {
              type: 'signin',
              value: `${process.env.DomainName}/bot-auth-start`
            }
          }
        }
      ]
    });
  }

  static simpleMessageCard(message) {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [{
          type: 'TextBlock',
          text: message,
          wrap: true
      }]
    });
  }

  static pleaseAddBotToChannelCard(conversationType) {
    let message;
    if (conversationType === 'groupChat')
      message = 'Please add the Quire bot to this conversation first.';
    else if (conversationType === 'channel')
      message = 'Please add the Quire bot to this channel first.'
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [{
          type: 'TextBlock',
          text: message,
          wrap: true
      }],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: 'Show me how',
          url: 'https://support.microsoft.com/en-us/office/add-an-app-to-microsoft-teams-b2217706-f7ed-4e64-8e96-c413afd02f77'
        }
      ]
    });
  }

  static unknownCommandCard() {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: 'Sorry, I am not quite sure what you mean, but I\'m here to help! Please use the below help button to see what I can do for you.',
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.Submit',
          title: 'Help',
          data: {
            msteams: {
              type: 'imBack',
              value: 'Help'
            }
          }
        }
      ]
    });
  }

  static unknownErrorCard() {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: 'Sorry, we encountered an unexpected error. We will look into it, '
            +', but feel free to contact us. Please use the below contact us button to see what I can do for you.',
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: 'Contact us',
          url: 'https://quire.io/feedback'
        }
      ]
    });
  }

  static taskCompleteCard(task) {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.2',
      body: [
        {
          type: 'TextBlock',
          text: `${task.nameText} has been completed.`,
          wrap: true
        }
      ],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: 'View in Quire',
          url: task.url
        }
      ]
    });
  }

  static helpCard() {
    return CardFactory.adaptiveCard({
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.3',
      body: [
        {
          type: 'TextBlock',
          text: 'Here are the commands you can use with this app:',
          wrap: true
        },
        {
          type: 'FactSet',
          facts: [
            {
              title: 'Add task',
              value: 'Add a new task in Quire'
            },
            {
              title: 'Link project',
              value: 'Link a project to this channel'
            },
            {
              title: 'Follow project',
              value: 'Follow a project'
            },
            {
              title: 'Unlink project',
              value: 'Unlink a project from this channel'
            },
            {
              title: 'Unfollow project',
              value: 'Get this channel to unfollow a project'
            },
            {
              title: 'Login',
              value: 'Log into Quire'
            },
            {
              title: 'Logout',
              value: 'Log out of Quire'
            },
            {
              title: 'Help',
              value: 'View a list of possible commands'
            }
          ]
        }
      ],
      actions: [
        {
          type: 'Action.OpenUrl',
          title: 'Learn more',
          url: 'https://quire.io/guide/microsoft-teams/'
        }
      ]
    });
  }

  static _tourCard(title, desc, image, buttons) {
    return CardFactory.adaptiveCard({
        type: 'AdaptiveCard',
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.3',
        body: [
          {
            type: 'TextBlock',
            size: 'large',
            weight: 'bolder',
            text: title,
          },
          {
            type: 'Image',
            url: image
          },
          {
            type: 'TextBlock',
            text: desc,
            wrap: true
          },
        ],
        actions: buttons
      });
  }

  static tourCard() {
    return [
      this._tourCard(
        'Welcome to Quire bot', 
        'You can interact with Quire bot from your team channel. ' +
            'Quire bot can help you add a task, assign and add comment for '+
            'a task via a series of actionable messages.',
        'https://d12y7sg0iam4lc.cloudfront.net/s/img/app/msteams/tour_1.png',
        [
          {
            type: 'Action.OpenUrl',
            title: 'Learn more',
            url: 'https://quire.io/guide/microsoft-teams/'
          }
        ]),
      this._tourCard(
        'Link a project in Quire', 
        'Link a specific project in Quire that you would want Microsoft Teams to access.',
        'https://d12y7sg0iam4lc.cloudfront.net/s/img/app/msteams/tour_2.png',
        [
          {
            type: 'Action.Submit',
            title: 'Link project',
            data: {
              fetchId: 'linkProject_fetch',
              msteams: {type: 'task/fetch'}
            }
          },
          {
            type: 'Action.OpenUrl',
            title: 'Learn more',
            url: 'https://quire.io/guide/microsoft-teams/'
          }
        ]),
      this._tourCard(
        'Add a new task to Quire', 
        'Once you have successfully linked a project, you can add your first task by sending a message '+
            '"Add task" and Quire will help you create your task automatically!',
        'https://d12y7sg0iam4lc.cloudfront.net/s/img/app/msteams/tour_3.png',
        [
          {
            type: 'Action.Submit',
            title: 'Add task',
            data: {
              fetchId: 'addTask_fetch',
              msteams: {type: 'task/fetch'}
            }
          },
          {
            type: 'Action.OpenUrl',
            title: 'Learn more',
            url: 'https://quire.io/guide/microsoft-teams/'
          }
        ]),
      this._tourCard(
        'Get help from Quire bot', 
        'At any time you want to look for help from Quire bot, just type "help" in the message composer. '+
            'The list of available commands that you can use with Quire and Microsoft Teams integration will be '+
            'presented for you.',
        'https://d12y7sg0iam4lc.cloudfront.net/s/img/app/msteams/tour_4.png',
        [
          {
            type: 'Action.Submit',
            title: 'Help',
            data: {
              msteams: {
                type: 'imBack',
                value: 'Help'
              }
            }
          },
          {
            type: 'Action.OpenUrl',
            title: 'Learn more',
            url: 'https://quire.io/guide/microsoft-teams/'
          }
        ]),
    ];
  }
}

module.exports.CardTemplates = CardTemplates;
