const { Signale } = require('signale');
const argv = require('minimist')(process.argv.slice(2));

const { name } = require('../package.json');

const logger = new Signale({
  interactive: false,
  scope: name,
  displayTimestamp: true,
  stream: process.stdout,
  logLevel: argv['sp-log-level'] || 'info',
  types: {
    progress: {
      badge: '✪',
      color: 'yellow',
      label: 'progress',
      logLevel: 'info',
    },
    complete: {
      badge: '√',
      color: 'green',
      label: 'complete',
      logLevel: 'info',
    },
    await: {
      badge: '●',
      color: 'gray',
      label: 'awaiting',
      logLevel: 'info',
    },
  },
});

module.exports = logger;
