// Copyright (C) 2020 Potix Corporation. All Rights Reserved
// History: 2020/12/22 5:55 PM
// Author: charlie<charliehsieh@potix.com>

const { QuireApi } = require("./quireApi");

let clientToken = {};

async function initClientToken() {
  const token = await QuireApi.getClientCredentialsToken();
  clientToken = token;
}

function getClientToken() {
  return clientToken;
}

async function getUserToken(teamsId) {
  return await QuireApi.getTokenFromStorage(teamsId);
}

async function isUserLogin(teamsId) {
  const userToken = await getUserToken(teamsId);
  return userToken.length != 0;
}

module.exports = {
  initClientToken: initClientToken,
  getClientToken: getClientToken,
  getUserToken: getUserToken,
  isUserLogin: isUserLogin
};
