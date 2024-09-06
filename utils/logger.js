
const winston = require('winston');

const logger = winston.createLogger({
  levels: winston.config.npm.levels,
  format: winston.format.combine(
    winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss:ms' }),
    winston.format.printf((info) => `[${info.timestamp}] ${info.level}: ${info.message}`)
  ),
  transports: [
    new winston.transports.Console(),
    new winston.transports.File({ filename: `${process.env.LOG_DIR}/error.log`, level: 'error' }),
    new winston.transports.File({ filename: `${process.env.LOG_DIR}/logs.log`}),
  ],
  exitOnError: false
})

module.exports = {
  logger: logger
}