const { logger, transformInput, isEmpty } = require('./utils');
const { setup } = require('./setup');
const { deploy } = require('./deploy');
const { accesstoken } = require('./accesstoken');

const main = async (input) => {
  const args = transformInput(input);
  try {
    if (!((!isEmpty(args.appClientId) && !isEmpty(args.appClientSecret)) || !isEmpty(args.accessToken))) {
      throw new Error('Sharepoint App Credentials or an AccessToken is required to access site'.red);
    } else if (isEmpty(args.siteUrl)) {
      throw new Error('Sharepoint site url is required to proceed'.red);
    } else if (args.accesstoken) {
      await accesstoken(args).catch((e) => {
        throw new Error((`Unable to get Access Token ${args.siteUrl}/${args.subsite}/\n`).red + e);
      });
    } else if (args.deploy) {
      await deploy(args).catch((e) => {
        throw new Error((`Unable to Deploy ${args.siteUrl}/${args.subsite}/\n`).red + e);
      });
    } else if (args.setup) {
      await setup(args).catch((e) => {
        throw new Error((`Unable to Setup ${args.siteUrl}/${args.subsite}/\n`).red + e);
      });
    } else {
      throw new Error('Plese use flags --deploy, --setup or --accesstoken to proceed'.red);
    }
  } catch (e) {
    logger.fatal(e);
    throw new Error(e);
  }
};

exports.main = main;
