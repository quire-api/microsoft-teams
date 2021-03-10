// Copyright (C) 2021 Potix Corporation. All Rights Reserved
// History: 2021/3/10 9:55 AM
// Author: jimmyshiau<jimmyshiau@potix.com>
const commandDescriptions = {
  'add task': 'adding a new task',
  'create task': 'adding a new task',
  'link project': 'linking a project',
  'follow project': 'following a project',
  'login': 'adding a new task',
};

const cardButtonLabels = {
  'taskComplete_submit': 'Complete task',
  'followTask_submit': 'Follow task',
  'addComment_submit': 'Add comment',
  'addComment_fetch': 'Add comment',
};

class QuireMessages {
    static getCommandDescriptions(command) {
        return commandDescriptions[command];
    }

    static getButtonLabel(actionId) {
        return cardButtonLabels[actionId];
    }
}

module.exports.QuireMessages = QuireMessages;