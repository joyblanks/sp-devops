const { logger, transformInput } = require('./utils');
const { setup } = require('./setup');
const { deploy } = require('./deploy');
const { accesstoken } = require('./accesstoken');

const main = async (input) => {
  const args = transformInput(input);
  try {
    if (!((args.appClientId && args.appClientSecret) || args.accessToken)) {
      logger.error('Sharepoint App Details required to access site');
      process.exit(0);
    } else if (!args.siteUrl) {
      logger.error('Sharepoint site is required to proceed');
      process.exit(0);
    } else if (args.accesstoken) {
      await accesstoken(args).catch((e) => {
        logger.error((`Unable to get Access Token ${args.site}/${args.subsite}/`).red);
        logger.error(e);
      });
    } else if (args.deploy) {
      await deploy(args).catch((e) => {
        logger.error((`Unable to Deploy ${args.site}/${args.subsite}/`).red);
        logger.error(e);
      });
    } else if (args.setup) {
      await setup(args).catch((e) => {
        logger.error((`Unable to Setup ${args.site}/${args.subsite}/`).red);
        logger.error(e);
      });
    } else {
      logger.fatal('Plese use flags --deploy, --setup or --accesstoken to proceed'.bgRed);
    }
  } catch (e) {
    const err = '[Error] Something went wrong';
    if (logger) {
      logger.error(err + e);
    } else {
      process.stderr.write(err + e);
    }
  }
};

exports.main = main;
