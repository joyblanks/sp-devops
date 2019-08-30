const fetch = require('node-fetch');
const path = require('path');
const fs = require('fs');
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

const readFile = (location) => fs.readFileSync(location);

const normalizeFilePathForUpload = (file) => file.split(path.sep).join('/');

const displayNoSubSite = (str, subsite) => {
  if (subsite) {
    return str.replace(new RegExp(`/${subsite}/`, 'g'), '');
  }
  return str.substring(1);
};

const transformInput = (input) => ({
  deploy: input.deploy,
  setup: input.setup,
  accesstoken: input.accesstoken,

  logLevel: input.SP_LOG_LEVEL,
  siteUrl: input.SP_SITE_URL,
  subsite: input.SP_SUBSITE,
  appClientId: input.SP_APP_CLIENT_ID,
  appClientSecret: input.SP_APP_CLIENT_SECRET,
  accessToken: input.SP_ACCESTOKEN,
  remoteFolder: input.SP_REMOTE_FOLDER,
  distFolder: input.SP_DIST_FOLDER,
  specFileCreateList: input.SP_SPEC_LIST || 'sharepoint-list-spec.json',
  specFileCreateSite: input.SP_SPEC_SITE || 'sharepoint-site-spec.json', // TODO
});

const throwError = (e) => {
  throw e;
};

const makeQuery = (params) => Object.keys(params).map((key) => `${key}=${encodeURIComponent(params[key])}`).join('&');

const extractClientInfo = (key, data) => {
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
    const errorMsg = response && response.error_description;
    if (errorMsg) {
      throwError(new Error(errorMsg));
    }
    response = response.d.GetContextWebInformation.FormDigestValue;
  } catch (e) {
    throwError(e);
  }
  return response;
};

const isEmpty = (val) => (val === undefined || val === null || val === '');

module.exports = {
  getFormDigestValue,
  transformInput,
  makePath,
  logger,
  getHeaders,
  colors,
  padGap,
  throwError,
  normalizeFilePathForUpload,
  displayNoSubSite,
  makeQuery,
  extractClientInfo,
  readFile,
  isEmpty,
};
