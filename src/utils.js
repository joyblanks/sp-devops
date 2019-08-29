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
  ...input,
  appClientId: input.sp_app_client_id || input['sp-app-client-id'],
  appClientSecret: input.sp_app_client_secret || input['sp-app-client-secret'],
  accessToken: input.sp_accestoken || input['sp-access-token'],
  siteUrl: input.sp_site_url || input['sp-site-url'],
  subsite: input.sp_subsite || input['sp-subsite'],
  remoteFolder: input.sp_remote_folder || input['sp-remote-folder'],
  distFolder: input.sp_dist_folder || input['sp-dist-folder'],
  specFileCreateList: input.sp_spec_list || input['sp-spec-list'] || 'sharepoint-list-spec.json',
  specFileCreateSite: input.sp_spec_site || input['sp-spec-site'] || 'sharepoint-site-spec.json', // TODO
});

const throwError = (e) => {
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
  extract,
  readFile,
  isEmpty,
};
