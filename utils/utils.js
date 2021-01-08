// Copyright (C) 2020 Potix Corporation. All Rights Reserved
// History: 2020/12/22 5:55 PM
// Author: charlie<charliehsieh@potix.com>

const dbAccess = require('../db/dbAccess');
const randomNumber = require("random-number-csprng");
const verificationCodeLength = 8;
const verificationCodeSurviveDuration = 60 * 1000; // 1 minute
const unsafeUserTokens = {};

async function prepareVerificationCode(token) {
  const verificationCode = await generateVerificationCode();
  unsafeUserTokens[verificationCode] = token;
  setTimeout(() => {
    delete unsafeUserTokens[verificationCode];
  }, verificationCodeSurviveDuration);
  return verificationCode;
}

async function generateVerificationCode() {
  let verificationCode = await randomNumber(0, Math.pow(10, verificationCodeLength) - 1);
  return ('0'.repeat(verificationCodeLength) + verificationCode).substr(-verificationCodeLength);
}

async function getUserTokenByVerificationCode(verificationCode) {
  const token = unsafeUserTokens[verificationCode];
  delete unsafeUserTokens[verificationCode];
  return token;
}

function addExpirationTimeForToken(token) {
  const durationInMillisecond = token.expires_in * 1000;
  token.expirationTime = durationInMillisecond + new Date().getTime();
}

function isTokenExpired(token) {
  return token.expirationTime < new Date();
}

async function isUserLogin(teamsId) {
  const token = await dbAccess.getToken(teamsId);
  return token ? true : false;
}

function getConversationId(activity) {
  const channelData = activity.channelData;
  const channel = channelData ? channelData.channel : null;
  if (channel) return channel.id;

  // sometimes conversation id will append with message id, e.g.
  // '19:78521d8e8608418180cf3fe89186ba87@thread.tacv2;messageid=1608194466346'
  // we only need the string before ';'
  const conversationId = activity.conversation.id;
  const idx = conversationId.indexOf(';');
  return idx == -1 ? conversationId : conversationId.substr(0, idx);
}

function itemsToChoices(array) {
  return array.map(elem => {
    return {
      title: elem.nameText,
      value: JSON.stringify({
        oid: elem.oid,
        nameText: elem.nameText
      })
    }
  }).sort((curr, next) => {
    if (curr.title > next.title) return 1;
    if (curr.title < next.title) return -1;
    return 0;
  });
}

function projectsToChoices(array) {
  let inbox;
  return array.map(elem => {
    if (elem.oid.substr(1) === elem.createdBy.oid) {
      elem.nameText = 'My tasks';
      inbox = elem;
    }

    return {
      oid: elem.oid,
      nameText: elem.nameText
    };
  }).sort((curr, next) => {
    if (curr.oid === inbox.oid) return -1;
    if (next.oid === inbox.oid) return 1;

    if (curr.nameText > next.nameText) return 1;
    if (curr.nameText < next.nameText) return -1;
    return 0;
  }).map(elem => {
    return {
      title: elem.nameText,
      value: JSON.stringify({
        oid: elem.oid,
        nameText: elem.nameText
      })
    };
  });
}

module.exports = {
  addExpirationTimeForToken: addExpirationTimeForToken,
  getUserTokenByVerificationCode: getUserTokenByVerificationCode,
  getConversationId: getConversationId,
  isTokenExpired: isTokenExpired,
  isUserLogin: isUserLogin,
  prepareVerificationCode: prepareVerificationCode,
  itemsToChoices: itemsToChoices,
  projectsToChoices: projectsToChoices
}
