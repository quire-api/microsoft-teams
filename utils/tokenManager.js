// Copyright (C) 2020 Potix Corporation. All Rights Reserved
// History: 2020/12/22 5:55 PM
// Author: charlie<charliehsieh@potix.com>

const { isTokenExpired, addExpirationTimeForToken } = require("./utils");
const { TeamsHttp } = require("./teamsHttp");

let clientToken = {};

async function initClientToken() {
  const token = await TeamsHttp.getClientCredentialsToken();
  addExpirationTimeForToken(token);
  clientToken = token;
}

async function getClientToken() {
  if (isTokenExpired(clientToken)) {
    await refreshClientToken();
  }
  return clientToken;
}

async function refreshClientToken() {
  const newToken = await TeamsHttp.refreshToken(clientToken);
  addExpirationTimeForToken(clientToken);
  clientToken = newToken;
}

async function getUserToken(teamsId) {
  let userToken = await TeamsHttp.getTokenFromStorage(teamsId);
  if (isTokenExpired(userToken)) {
    userToken = await TeamsHttp.refreshToken(userToken);
    addExpirationTimeForToken(userToken);
    await TeamsHttp.putTokenToStorage(teamsId, userToken);
  }
  return userToken;
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
