// Copyright (C) 2020 Potix Corporation. All Rights Reserved
// History: 2020/12/22 5:55 PM
// Author: charlie<charliehsieh@potix.com>

const axios = require('axios');
const dbAccess = require('../db/dbAccess');
const querystring = require('querystring');
const utils = require('./utils');
const clientId = process.env.QuireAppId;
const clientSecret = process.env.QuireAppSecret;
const domainName = process.env.DomainName;
const quireUrl = process.env.QuireUrl;
const tokenUrl = `${quireUrl}/oauth/token`;
const apiUrl = `${quireUrl}/api`;

class QuireApi {

  static async getClientCredentialsToken() {
    return axios.post(tokenUrl, querystring.encode({
      grant_type: 'client_credentials',
      client_id: clientId,
      client_secret: clientSecret
    })).then(res => res.data);
  }

  static async _refreshToken(oldToken) {
    return axios.post(tokenUrl, querystring.encode({
      grant_type: 'refresh_token',
      refresh_token: oldToken.refresh_token,
      client_id: clientId,
      client_secret: clientSecret
    })).then(res => res.data)
    .catch(error => {
      // if refresh failed, Quire will return 400
      if (error.response.status === 400) {
        return {isInvalidToken: true};
      }
      throw error;
    });
  }

  static async refreshAndStoreToken(teamsId, oldToken) {
    const newToken = await this._refreshToken(oldToken);
    if (newToken.isInvalidToken) {
      dbAccess.deleteToken(teamsId);
    } else {
      dbAccess.putToken(teamsId, newToken);
    }
    return newToken;
  }

  // GET /user/list
  static async getAllUsers(token) {
    return axios.get(`${apiUrl}/user/list`, authHeader(token))
    .then(res => res.data);
  }

  static async getCurrentUser(token) {
    return axios.get(`${apiUrl}/user/list`, authHeader(token))
    .then(res => [res.data[0]]);
  }

  // TODO must handle My tasks(oid = '-')
  // GET /user/list/project/{oid}
  static async getUsersByProjectOid(token, oid) {
    if (oid === '-') return await this.getCurrentUser(token);
    return axios.get(`${apiUrl}/user/list/project/${oid}`, authHeader(token))
    .then(res => res.data);
  }

  // GET /project/list params: { 'add-task' = true }
  static async getAllProjects(token) {
    return axios.get(`${apiUrl}/project/list?add-task=true`, authHeader(token))
    .then(res => res.data);
  }

  // GET /project/{oid}
  static async getProjectByOid(token, oid) {
    return axios.get(`${apiUrl}/project/${oid}`, authHeader(token))
    .then(res => res.data);
  }

  // POST /task/{oid}
  static async addTaskToProjectByOid(token, task, oid) {
    return axios.post(`${apiUrl}/task/${oid}`, task, authHeader(token))
    .then(res => res.data);
  }

  // POST /comment/{oid}
  static async addCommentToTaskByOid(token, comment, oid) {
    return axios.post(`${apiUrl}/comment/${oid}`, {
      description: comment
    }, authHeader(token)).then(res => res.data);
  }

  // PUT /task/{oid}
  static async setTaskComplete(token, oid) {
    return axios.put(`${apiUrl}/task/${oid}`, {
      status: 100
    }, authHeader(token)).then(res => res.data);
  }

  // GET /task/search/{oid}
  static async searchTaskByProjectOid(token, text, oid) {
    return axios.get(`${apiUrl}/task/search/${oid}`, {
      ...authHeader(token),
      params: {text: text}
    }).then(res => res.data);
  }

  // GET /task/{oid}
  static async getTaskByOid(token, oid) {
    return axios(`${apiUrl}/task/${oid}`, authHeader(token))
    .then(res => res.data);
  }

  // GET /task/list/{oid}
  static async getRootTasksByOid(token, oid) {
    return axios(`${apiUrl}/task/list/${oid}`, authHeader(token))
    .then(res => res.data);
  }

  // PUT /project/{oid}
  // syntax { "addFollowers": ["app|/path|channel"]}
  static async addFollowerToProject(token, projectOid, conversationId, serviceUrl) {
    return axios.put(`${apiUrl}/project/${projectOid}`, {
      addFollowers: [`app|/${conversationId}|${serviceUrl}`]
    }, authHeader(token))
    .then(res => res.data);
  }

  // PUT /project/{oid}
  // syntax { "removeFollowers": ["app|/path|channel"]}
  static async removeFollowerFromProject(token, projectOid, conversationId, serviceUrl) {
    return axios.put(`${apiUrl}/project/${projectOid}`, {
      removeFollowers: [`app|/${conversationId}|${serviceUrl}`]
    }, authHeader(token))
    .then(res => res.data);
  }

  // PUT /task/{oid}
  // syntax { "addFollowers": ["app|/path|channel"]}
  static async addFollowerToTask(token, taskOid, conversationId, serviceUrl) {
    return axios.put(`${apiUrl}/task/${taskOid}`, {
      addFollowers: [`app|/${conversationId}|${serviceUrl}`]
    }, authHeader(token))
    .then(res => res.data);
  }

  static async handleAuthStart(req, res) {
    const redirectUri = encodeURIComponent(`https://${domainName}/bot-auth-end`);
    const encodedId = encodeURIComponent(clientId);
    const url = `${quireUrl}/oauth?client_id=${encodedId}&redirect_uri=${redirectUri}`;
    let resBody = '<html><head><title>Sign In</title></head><body>';
    resBody += '<script src="https://statics.teams.cdn.office.net/sdk/v1.6.0/js/MicrosoftTeams.min.js" integrity="sha384-mhp2E+BLMiZLe7rDIzj19WjgXJeI32NkPvrvvZBrMi5IvWup/1NUfS5xuYN5S3VT" crossorigin="anonymous"></script>';
    resBody += '<script type="text/javascript">';
    resBody += 'microsoftTeams.initialize();';
    resBody += `window.location.assign('${url}');`;
    resBody += '</script></body></html>';
    res.send(resBody);
  }

  static async handleAuthEnd(req, res) {
    const code = req.query.code;
    const postRes = await axios.post(tokenUrl, querystring.encode({
      grant_type: 'authorization_code',
      code: code,
      client_id: clientId,
      client_secret: clientSecret
    }));

    let resBody = '<html><head><title>Quire for Teams Authentication</title></head>';
    resBody += '<body>';
    resBody += '<script src="https://statics.teams.cdn.office.net/sdk/v1.6.0/js/MicrosoftTeams.min.js" integrity="sha384-mhp2E+BLMiZLe7rDIzj19WjgXJeI32NkPvrvvZBrMi5IvWup/1NUfS5xuYN5S3VT" crossorigin="anonymous"></script>';
    resBody += '<script type="text/javascript">';
    resBody += 'microsoftTeams.initialize();';

    if (postRes.status == 200) {
      const verificationCode = await utils.prepareVerificationCode(postRes.data);
      resBody += `microsoftTeams.authentication.notifySuccess('${verificationCode}');`;
    } else {
      resBody += 'microsoftTeams.authentication.notifyFailure();';
    }

    resBody += '</script></body></html>';
    res.send(resBody);
  }
}

function authHeader(token) {
  const header = { headers: { 'Authorization': `Bearer ${token.access_token}` } };
  return header;
}

module.exports.QuireApi = QuireApi;