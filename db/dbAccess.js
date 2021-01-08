// Copyright (C) 2021 Potix Corporation. All Rights Reserved
// History: 2021/1/8 12:03 PM
// Author: charlie<charliehsieh@potix.com>

const sqlite3 = require('sqlite3');
const db = new sqlite3.Database(process.env.DBPath);

const tokenDuration = 180 * 24 * 60 * 60 * 1000; // 6 months
const clearInterval = 24 * 60 * 60 * 1000; // 24 hours
let lastClearTime = 0;

function _clearToken() {
  const now = new Date().getTime();
  if (now - lastClearTime > clearInterval) {
    lastClearTime = now;
    const time = now - tokenDuration;
    db.run(`DELETE FROM token WHERE lastAccessTime < ?`, time);
  }
}

function putToken(teamsId, token) {
  db.get(`SELECT teamsId FROM token WHERE teamsId = ?`, teamsId, (err, row) => {
    let sql;
    if (row) {
      sql = `UPDATE token SET accessToken = $accessToken,
          refreshToken = $refreshToken, lastAccessTime = $lastAccessTime
          WHERE teamsId = $teamsId`;
    } else {
      sql = `INSERT INTO token VALUES ($teamsId, $accessToken, $refreshToken, $lastAccessTime)`;
    }
    db.run(sql, {
      $teamsId: teamsId,
      $accessToken: token.access_token,
      $refreshToken: token.refresh_token,
      $lastAccessTime: new Date().getTime()
    });
  });
}

function getToken(teamsId) {
  _clearToken();
  return new Promise((resolve, reject) => {
    db.get(`SELECT * FROM token WHERE teamsId = ?`, teamsId, (err, row) => {
      const token = row ? {
        access_token: row.accessToken,
        refresh_token: row.refreshToken
      } : undefined;
      resolve(token);

      if (token) {
        db.run(`UPDATE token SET lastAccessTime = ? WHERE teamsId = ?`,
            new Date().getTime(), teamsId);
      }
    });
  });
}

function deleteToken(teamsId) {
  db.run(`DELETE FROM token WHERE teamsId = ?`, teamsId);
}

function putLinkedProject(id, project) {
  db.get(`SELECT id FROM linkedProject WHERE id = ?`, id ,(err, row) => {
    let sql;
    if (row) {
      sql = `UPDATE linkedProject
          SET oid = $oid, nameText = $nameText, lastAccessTime = $lastAccessTime
          WHERE id = $id`;
    } else {
      sql = `INSERT INTO linkedProject VALUES ($id, $oid, $nameText, $lastAccessTime)`;
    }
    db.run(sql, {
      $id: id,
      $oid: project.oid,
      $nameText: project.nameText,
      $lastAccessTime: new Date().getTime()
    });
  });
}

function getLinkedProject(id) {
  return new Promise((resolve, reject) => {
    db.get(`SELECT * FROM linkedProject WHERE id = ?`, id, (err, row) => {
      resolve(row);

      if (row) {
        db.run(`UPDATE linkedProject SET lastAccessTime = ?
            WHERE id = ?`, new Date().getTime(), id);
      }
    });
  });
}

function deleteLinkedProject(id) {
  db.run(`DELETE FROM linkedProject WHERE id = ?`, id);
}

module.exports = {
  putToken: putToken,
  getToken: getToken,
  deleteToken: deleteToken,
  putLinkedProject: putLinkedProject,
  getLinkedProject: getLinkedProject,
  deleteLinkedProject: deleteLinkedProject
}
