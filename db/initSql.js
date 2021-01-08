const sqlite3 = require('sqlite3');
const db = new sqlite3.Database('storage.db');
 
db.serialize(function() {
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
  )`)
});
 
db.close();
