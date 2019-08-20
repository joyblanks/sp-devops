const https = require('https');

const {
  makePath,
  logger,
  getHeaders,
  padGap,
  throwError,
  extract,
  makeQuery,
} = require('./utils');

const getClientInfo = async (site, subsite) => new Promise((resolve) => {
  const req = https.get({
    hostname: new URL(site).hostname,
    path: `${makePath(null, subsite)}/_vti_bin/client.svc`,
    headers: {
      Authorization: 'Bearer',
    },
  }, (res) => {
    const authHeader = res.headers['www-authenticate'];
    const realm = extract('realm', authHeader);
    const clientId = extract('client_id', authHeader);
    resolve({ realm, clientId });
  });
  req.on('error', throwError);
  req.end();
}).catch(throwError);

const getAccessToken = async (site, subsite, appClientId, appClientSecret) => {
  const { realm, clientId } = await getClientInfo(site, subsite);
  logger.info('Generating'.padEnd(padGap), `${'AccessToken'.blue} App[${appClientId.magenta}@${realm.magenta}]`);
  const query = makeQuery({
    grant_type: 'client_credentials',
    client_id: `${appClientId}@${realm}`,
    client_secret: appClientSecret,
    resource: `${clientId}/${new URL(site).hostname}@${realm}`,
  });
  const options = {
    method: 'GET',
    hostname: 'accounts.accesscontrol.windows.net',
    path: `/${realm}/tokens/OAuth/2`,
    headers: getHeaders(null, {
      Host: 'accounts.accesscontrol.windows.net',
      'Content-Type': 'application/x-www-form-urlencoded',
      'Content-Length': query.length,
    }, ['Authorization', 'Accept']),
  };
  return new Promise((resolve) => {
    const req = https.request(options, (res) => {
      const chunks = [];
      res.on('data', (chunk) => {
        chunks.push(chunk);
      });
      res.on('end', () => {
        const body = Buffer.concat(chunks);
        try {
          const json = JSON.parse(body.toString());
          resolve(`Bearer ${json.access_token}`);
        } catch (e) {
          throwError(e);
        }
      });
    }).on('error', throwError);
    req.write(query);
    req.end();
  }).catch(throwError);
};

const printable = (json) => {
  const string = JSON.stringify(json);
  const token = json.Authorization.replace('Bearer ', '');
  return string
    .replace('Authorization', 'Authorization'.blue)
    .replace('":"', `" ${':'.green} "`)
    .replace(/"/g, '"'.yellow)
    .replace('Bearer', 'Bearer'.grey)
    .replace('{', '{ '.cyan)
    .replace('}', ' }'.cyan)
    .replace(token, token.green);
};

const accesstoken = async ({
  appClientId,
  appClientSecret,
  siteUrl,
  subsite,
}) => {
  logger.start('Operation Started'.bold, '😈');
  try {
    const authorization = await getAccessToken(siteUrl, subsite, appClientId, appClientSecret).catch(throwError);
    logger.complete('AccessToken', '\n', printable({ Authorization: authorization }));
    logger.success('Operation Completed'.bold, '😈');
  } catch (e) {
    throwError(e);
  }
};

module.exports = {
  getAccessToken,
  accesstoken,
};
