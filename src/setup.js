const fetch = require('node-fetch');
const fs = require('fs');

const {
  getFormDigestValue,
  makePath,
  logger,
  getHeaders,
  padGap,
  throwError,
} = require('./utils');
const { getAccessToken } = require('./accesstoken');


const isListExists = async (siteUrl, subsite, listName, authorization) => {
  let response = null;
  try {
    response = await fetch(`${makePath(siteUrl, subsite)}/_api/web/lists/getbytitle('${listName}')`, {
      headers: getHeaders(authorization),
      method: 'GET',
    }).catch(throwError);
    response = await response.json().catch(throwError);
    response = !!(response.d);
  } catch (e) {
    throwError(e);
  }
  return response;
};

const getListColumns = async (siteUrl, subsite, listName, authorization) => {
  let response = null;
  const query = '$filter=Hidden eq false and ReadOnlyField eq false&$select=Title';
  try {
    response = await fetch(`${makePath(siteUrl, subsite)}/_api/web/lists/getbytitle('${listName}')/Fields?${query}`, {
      headers: getHeaders(authorization),
      method: 'GET',
    }).catch(throwError);
    response = await response.json().catch(throwError);
    if (response.d && response.d.results) {
      response = response.d.results.map((col) => col.Title);
    } else response = [];
  } catch (e) {
    throwError(e);
  }
  return response;
};

const deleteItems = async (siteUrl, subsite, listName, itemIds, authorization, formDigest) => {
  let count = 0;
  const total = itemIds.length;
  try {
    for (const id of itemIds) {
      logger.progress('Deleteing Item'.padEnd(padGap), (`${listName}(${++count}/${total})`).grey);
      await fetch(`${makePath(siteUrl, subsite)}/_api/web/lists/getbytitle('${listName}')/Items(${id})`, {
        body: null,
        headers: getHeaders(authorization, {
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*',
          'X-RequestDigest': formDigest,
        }),
        method: 'POST',
      }).catch(throwError);
    }
  } catch (e) {
    throwError(e);
  }
  return count;
  // TODO : convert to batch
};

const getItemIds = async (siteUrl, subsite, listName, authorization, next) => {
  const size = 1000;
  let ids = [];
  const url = `${makePath(siteUrl, subsite)}/_api/web/lists/getbytitle('${listName}')/Items?$top=${size}&$select=Id`;
  try {
    let response = await fetch(next || url, {
      headers: getHeaders(authorization),
      method: 'GET',
    }).catch(throwError);
    response = await response.json().catch(throwError);
    ids = (response.d.results || []).map((item) => item.Id);
    if (response['odata.nextLink']) {
      return ids.concat(
        await getItemIds(siteUrl, subsite, listName, authorization, response['odata.nextLink']).catch(throwError),
      );
    }
  } catch (e) {
    throwError(e);
  }
  return ids;
};

const deleteList = async (siteUrl, subsite, listName, authorization, formDigest) => {
  logger.info('Deleting List'.padEnd(padGap), listName.red);
  let response = null;
  try {
    const itemIds = await getItemIds(siteUrl, subsite, listName, authorization).catch(throwError);
    await deleteItems(siteUrl, subsite, listName, itemIds, authorization, formDigest).catch(throwError);
    response = await fetch(`${makePath(siteUrl, subsite)}/_api/web/lists/getbytitle('${listName}')`, {
      body: null,
      headers: getHeaders(authorization, {
        'X-HTTP-Method': 'DELETE',
        'If-Match': '*',
        'X-RequestDigest': formDigest,
      }),
      method: 'POST',
    }).catch(throwError);
  } catch (e) {
    throwError(e);
  }
  return response;
};

const createList = async (siteUrl, subsite, listConfig, authorization, formDigest) => {
  logger.info('Creating List'.padEnd(padGap), listConfig.name.green);
  let response = null;
  try {
    response = await fetch(`${makePath(siteUrl, subsite)}/_api/web/lists`, {
      body: JSON.stringify({
        __metadata: { type: 'SP.List' },
        BaseTemplate: 100,
        Title: listConfig.name,
      }),
      headers: getHeaders(authorization, { 'X-RequestDigest': formDigest }),
      method: 'POST',
    }).catch(throwError);
    response = await response.json().catch(throwError);

    if (listConfig.addToQuickLaunch) {
      await fetch(`${makePath(siteUrl, subsite)}/_api/web/navigation/QuickLaunch`, {
        body: JSON.stringify({
          __metadata: { type: 'SP.NavigationNode' },
          Title: listConfig.name,
          Url: `${makePath(siteUrl, subsite)}/Lists/${listConfig.name}/AllItems.aspx`,
        }),
        headers: getHeaders(authorization, { 'X-RequestDigest': formDigest }),
        method: 'POST',
      }).catch(throwError);
    }
  } catch (e) {
    throwError(e);
  }
  return response;
};

const createColumn = async (siteUrl, subsite, listConfig, column, authorization, formDigest) => {
  logger.info('Creating Column'.padEnd(padGap), (`${listConfig.name}.${column.Title}`).green);
  let response = null;
  try {
    response = await fetch(`${makePath(siteUrl, subsite)}/_api/web/lists/getbytitle('${listConfig.name}')/fields`, {
      body: JSON.stringify(column),
      headers: getHeaders(authorization, { 'X-RequestDigest': formDigest }),
      method: 'POST',
    }).catch(throwError);
    response = await response.json().catch(throwError);

    if (listConfig.addToView) {
      const viewUrl = `Views/GetByTitle('All%20Items')/ViewFields/addViewField('${column.Title}')`;
      await fetch(`${makePath(siteUrl, subsite)}/_api/web/lists/getbytitle('${listConfig.name}')/${viewUrl}`, {
        body: null,
        headers: getHeaders(authorization, { 'X-RequestDigest': formDigest }),
        method: 'POST',
      }).catch(throwError);
    }
  } catch (e) {
    throwError(e);
  }
  return response;
};

const createColumns = async (siteUrl, subsite, listConfig, existingCols, authorization, formDigest) => {
  const responses = [];
  let response = null;
  try {
    for (const column of listConfig.columns) {
      if (!existingCols.includes(column.Title)) {
        response = await createColumn(siteUrl, subsite, listConfig, column, authorization, formDigest).catch(throwError);
        responses.push(response);
      }
    }
  } catch (e) {
    throwError(e);
  }
  return responses;
};

const addItems = async (siteUrl, subsite, listName, data = [], authorization, formDigest) => {
  let response = null;
  const responses = [];
  let count = 0;
  try {
    for (const item of data) {
      logger.progress('Populating Data'.padEnd(padGap), `${listName}[${item.Title || (`Item${++count}`)}]`);
      response = await fetch(`${makePath(siteUrl, subsite)}/_api/web/lists/getbytitle('${listName}')/Items`, {
        body: JSON.stringify({ ...item, __metadata: { type: `SP.Data.${listName}ListItem` } }),
        headers: getHeaders(authorization, { 'X-RequestDigest': formDigest }),
        method: 'POST',
      }).catch(throwError);
      response = await response.json().catch(throwError);
      responses.push(response);
    }
  } catch (e) {
    throwError(e);
  }
  return responses;
};

const createLists = async (siteUrl, subsite, specFileCreateList, authorization, formDigest) => {
  try {
    const rawData = fs.readFileSync(specFileCreateList);
    const specJson = JSON.parse(rawData);
    for (const listConfig of specJson.config) {
      const listExists = await isListExists(siteUrl, subsite, listConfig.name, authorization).catch(throwError);
      let existingColumns = [];
      if (listExists) {
        if (listConfig.dropIfExists) {
          await deleteList(siteUrl, subsite, listConfig.name, authorization, formDigest).catch(throwError);
          await createList(siteUrl, subsite, listConfig, authorization, formDigest).catch(throwError);
        } else {
          existingColumns = await getListColumns(siteUrl, subsite, listConfig.name, authorization).catch(throwError);
        }
      } else {
        await createList(siteUrl, subsite, listConfig, authorization, formDigest).catch(throwError);
      }
      await createColumns(siteUrl, subsite, listConfig, existingColumns, authorization, formDigest).catch(throwError);
      await addItems(siteUrl, subsite, listConfig.name, listConfig.items, authorization, formDigest).catch(throwError);
    }
  } catch (e) {
    throwError(e);
  }
};

const setup = async ({
  appClientId,
  appClientSecret,
  siteUrl,
  subsite,
  specFileCreateList,
  authToken,
}) => {
  logger.start('Operation Started'.bold, '😈');
  try {
    const authorization = authToken || await getAccessToken(siteUrl, subsite, appClientId, appClientSecret).catch(throwError);
    const formDigest = await getFormDigestValue(siteUrl, subsite, authorization).catch(throwError);
    await createLists(siteUrl, subsite, specFileCreateList, authorization, formDigest).catch(throwError);
    logger.success('Operation Completed'.bold, '😈');
  } catch (e) {
    throwError(e);
  }
};

exports.setup = setup;
