// Copyright (C) 2021 Potix Corporation. All Rights Reserved
// History: 2021/1/14 11:03 AM
// Author: charlie<charliehsieh@potix.com>

const path = require('path');
const ENV_FILE = path.join(__dirname, '../../boeneo/mis/configurations/config/msteams/env');
require('dotenv').config({ path: ENV_FILE });
const sqlite3 = require('sqlite3');
const db = new sqlite3.Database(process.env.DBPath);
 
db.serialize(() => {
  db.run(`DROP TABLE token`, () => {});
  db.run(`DROP TABLE linkedProject`, () => {});

  db.run(`CREATE TABLE token (
    teamsId TEXT PRIMARY KEY NOT NULL,
    accessToken TEXT NOT NULL,
    refreshToken TEXT NOT NULL,
    lastAccessTime INTEGER NOT NULL
  )`);

  db.run(`CREATE TABLE linkedProject (
    id TEXT PRIMARY KEY NOT NULL,
    oid TEXT NOT NULL,
    nameText TEXT NOT NULL,
    lastAccessTime INTEGER NOT NULL
  )`);
});
 
db.close();
