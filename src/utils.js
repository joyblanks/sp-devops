const https = require('https');
const fetch = require('node-fetch');
const path = require('path');
const colors = require('colors');
const logger = require('./logger');


const padGap = 20;

const makePath = (site, subsite = '') => {
  if (site && subsite) {
    return `${site}/${subsite}`;
  }
  if (!site && subsite) {
    return `/${subsite}`;
  }
  if (site && !subsite) {
    return site;
  }
  return '';
};

const normalizeFilePathForUpload = (file) => file.split(path.sep).join('/');

const displayNoSubSite = (str, subsite) => {
  if (subsite) {
    return str.replace(new RegExp(`/${subsite}/`, 'g'), '');
  }
  return str.substring(1);
};

const transformInput = (input) => ({
  ...input,
  appClientId: input['sp-app-client-id'],
  appClientSecret: input['sp-app-client-secret'],
  accessToken: input['sp-access-token'],
  siteUrl: input['sp-site-url'],
  subsite: input['sp-subsite'],
  remoteFolder: input['sp-remote-folder'],
  distFolder: input['sp-dist-folder'],
  specFileCreateList: input['sp-spec-list'] || 'sharepoint-list-spec.json',
  specFileCreateSite: input['sp-spec-site'] || 'sharepoint-site-spec.json', // TODO
});

const throwError = (e) => {
  // logger.error(e);
  throw e;
};

const makeQuery = (params) => Object.keys(params).map((key) => `${key}=${encodeURIComponent(params[key])}`).join('&');

const extract = (key, data) => {
  const regex = new RegExp(`${key}="(.*?)"`);
  const values = regex.exec(data);
  return values && values[1];
};

const getHeaders = (authorization, headers = {}, removeList = []) => {
  const newHeaders = {
    Accept: 'application/json;odata=verbose',
    'Content-Type': 'application/json;odata=verbose',
    ...headers,
    Authorization: authorization,
  };
  for (const header of removeList) {
    delete newHeaders[header];
  }
  return newHeaders;
};

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
  logger.info('Generating'.padEnd(padGap), `${'AccessToken'.blue}=App[${appClientId.magenta}@${realm.magenta}]`);
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
    headers: {
      Host: 'accounts.accesscontrol.windows.net',
      'Content-Type': 'application/x-www-form-urlencoded',
      'Content-Length': query.length,
    },
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

const getFormDigestValue = async (site, subsite, authorization) => {
  logger.info('Generating'.padEnd(padGap), 'ContextInfo.FormDigest'.blue);
  let response = null;
  try {
    response = await fetch(`${makePath(site, subsite)}/_api/contextinfo`, {
      method: 'POST',
      body: null,
      headers: getHeaders(authorization),
    });
    response = await response.json();
    response = response.d.GetContextWebInformation.FormDigestValue;
  } catch (e) {
    logger.info(e);
    throw e;
  }
  return response;
};

module.exports = {
  getFormDigestValue,
  getAccessToken,
  transformInput,
  makePath,
  logger,
  getHeaders,
  colors,
  padGap,
  throwError,
  normalizeFilePathForUpload,
  displayNoSubSite,
};
