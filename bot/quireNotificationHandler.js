// Copyright (C) 2020 Potix Corporation. All Rights Reserved
// History: 2020/12/22 5:55 PM
// Author: charlie<charliehsieh@potix.com>

const { MessageFactory, CardFactory } = require("botbuilder");
const { CardTemplates } = require("../model/cardtemplates");
const { TeamsHttp } = require("../utils/teamsHttp");

const notificationType = {
  AddTask: 0,
  RemoveTask: 1,
  EditTask: 3,
  MoveTask: 4,
  Complete: 5,
  Uncomplete: 6,
  Assign: 7,
  Unassign: 8,
  SetDue: 9,
  UnsetDue: 10,
  SetState: 11,
  AddTaskComment: 16,
  AddTaskAttachment: 20,
  RemoveTaskAttachment: 21,
  SetTag: 28,
  UnsetTag: 29,
  SetPriority: 35,
  SetStart: 38,
  UnsetStart: 39,
  RemindStart: 80,
  RemindDue: 81,
  RemindOverdue: 82,
  SetBoard: 85,
  UnsetBoard: 86,
  AddProject: 100,
  RemoveProject: 101,
  AddProjectMember: 104,
  RemoveProjectMember: 105,
  AddProjectComment: 109,
  AddProjectAttachment: 110
}

async function handleQuireNotification(context, data) {
  switch (data.type) {
    case notificationType.AddTask:
      const task = await TeamsHttp.getTaskByOid(data.what.oid);
      let taskCard = CardTemplates.addHeaderToCard(CardTemplates.taskCard(task), data.text);
      taskCard = CardFactory.adaptiveCard(taskCard);
      await context.sendActivity(MessageFactory.attachment(taskCard));
      break;
    default:
      const msg = MessageFactory.text(data.message);
      await context.sendActivity(msg);
  }
}

module.exports = {
  handleQuireNotification: handleQuireNotification
}
