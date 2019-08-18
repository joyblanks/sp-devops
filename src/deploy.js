const fetch = require('node-fetch');
const fs = require('fs');
const path = require('path');
const { promisify } = require('util');
const {
  getFormDigestValue,
  getAccessToken,
  makePath,
  logger,
  getHeaders,
  padGap,
  throwError,
  normalizeFilePathForUpload,
  displayNoSubSite,
} = require('./utils');

const getFiles = async (siteUrl, subsite, folder, authorization) => {
  logger.info('Listing Folder'.padEnd(padGap), folder.cyan);
  let response = null;
  try {
    response = await fetch(`${makePath(siteUrl, subsite)}/_api/web/getFolderByServerRelativeUrl('${folder}')/files`, {
      method: 'GET',
      body: null,
      headers: getHeaders(authorization),
    }).catch(throwError);
    response = await response.json().catch(throwError);
    response = response.d.results.map((file) => file.ServerRelativeUrl);
  } catch (e) {
    throwError(e);
  }
  return response;
};

const deleteFile = async (siteUrl, subsite, file, formDigest, authorization) => {
  logger.progress('Deleting File'.padEnd(padGap), displayNoSubSite(file, subsite).red);
  let response = null;
  try {
    response = await fetch(`${makePath(siteUrl, subsite)}/_api/web/GetFileByServerRelativeUrl('${file}')`, {
      method: 'POST',
      body: null,
      headers: getHeaders(authorization, {
        'X-HTTP-Method': 'DELETE',
        'If-Match': '*',
        'X-RequestDigest': formDigest,
      }),
    }).catch(throwError);
    // response = await response.text().catch(throwError); // There is no response after delete
  } catch (e) {
    throwError(e);
  }
  return response;
};

const deleteFolder = async (siteUrl, subsite, folder, formDigest, authorization) => {
  logger.progress('Deleting Folder'.padEnd(padGap), (`[${displayNoSubSite(folder, subsite)}]`).red.bold);
  let response = null;
  try {
    response = await fetch(`${makePath(siteUrl, subsite)}/_api/web/GetFolderByServerRelativeUrl('${folder}')`, {
      method: 'POST',
      body: null,
      headers: getHeaders(authorization, {
        'X-HTTP-Method': 'DELETE',
        'If-Match': '*',
        'X-RequestDigest': formDigest,
      }),
    }).catch(throwError);
    // response = await response.json();
  } catch (e) {
    throwError(e);
  }
  return response;
};

const makeFolder = async (siteUrl, subsite, folder, remoteFolder, formDigest, authorization) => {
  let spFolder = remoteFolder;
  const responseArray = [];
  try {
    const folderArray = folder.split('/');
    while (folderArray.length) {
      const sub = folderArray.shift();
      logger.progress('Creating folder'.padEnd(padGap), (`[${spFolder}${sub ? '/' : ''}${sub}]`).green.bold);
      const object = JSON.stringify({ __metadata: { type: 'SP.Folder' }, ServerRelativeUrl: `${spFolder}/${sub}` });
      let response = await fetch(`${makePath(siteUrl, subsite)}/_api/web/folders`, {
        method: 'POST',
        body: object,
        headers: getHeaders(authorization, {
          'Content-Length': object.length,
          'X-RequestDigest': formDigest,
        }),
      }).catch(throwError);
      response = await response.text().catch(throwError);
      responseArray.push(response);
      spFolder = `${spFolder}/${sub}`;
    }
  } catch (e) {
    throwError(e);
  }
  return responseArray;
};

const getFolderHierarchy = async (siteUrl, subsite, folder, authorization) => {
  logger.info('Fetching Hierarchy'.padEnd(padGap), folder.cyan);
  const folderPath = `('${makePath(null, subsite)}/${folder}')?$expand=Folders,Files`;
  let response = null;
  try {
    response = await fetch(`${makePath(siteUrl, subsite)}/_api/Web/GetFolderByServerRelativeUrl${folderPath}`, {
      method: 'GET',
      body: null,
      headers: getHeaders(authorization),
    }).catch(throwError);
    let folders = [folder];
    response = await response.json().catch(throwError);
    if (response.d.Folders.results.length) {
      for (const subFolder of response.d.Folders.results) {
        if (subFolder.Name !== 'Forms') { // Ignore forms
          folders.push(`${folder}/${subFolder.Name}`);
          folders = folders.concat(
            await getFolderHierarchy(siteUrl, subsite, `${folder}/${subFolder.Name}`, authorization)
              .catch(throwError),
          );
        }
      }
    }
    response = [...new Set(folders)];
  } catch (e) {
    throwError(e);
  }
  return response;
};

const deleteSite = async (siteUrl, subsite, remoteFolder, formDigest, authorization) => {
  logger.await('Deleting Site'.padEnd(padGap), (`${makePath(siteUrl, subsite)}/${remoteFolder}`).red);
  try {
    const folders = await getFolderHierarchy(siteUrl, subsite, remoteFolder, authorization).catch(throwError);
    for (const folder of folders) {
      const files = await getFiles(siteUrl, subsite, folder, authorization).catch(throwError);
      for (const file of files) {
        await deleteFile(siteUrl, subsite, file, formDigest, authorization).catch(throwError);
      }
      if (folder !== remoteFolder) {
        await deleteFolder(siteUrl, subsite, `${makePath(null, subsite)}/${folder}`, formDigest, authorization)
          .catch(throwError);
      }
    }
    logger.complete('Deleted Site'.padEnd(padGap), (`${makePath(siteUrl, subsite)}/${remoteFolder}`.red));
  } catch (e) {
    throwError(e);
  }
};

const readFile = (location) => fs.readFileSync(location);

const listLocalFiles = async (dir, subdir) => {
  const readdirAsync = promisify(fs.readdir);
  const dirs = [];
  let filesInFolder = [];
  let response = null;
  try {
    const files = await readdirAsync(dir + path.sep + subdir).catch(throwError);
    for (const file of files) {
      if (fs.statSync(`${dir}${path.sep}${subdir}${path.sep}${file}`).isDirectory()) {
        filesInFolder = await listLocalFiles(`${dir}`, `${subdir}${path.sep}${file}`).catch(throwError);
        dirs.push(file);
      }
    }
    dirs.forEach((file) => {
      const index = files.indexOf(file);
      if (index !== -1) {
        files.splice(index, 1);
      }
    });
    response = files.map((file) => `${dir}${subdir}${path.sep}${file}`).concat(filesInFolder);
  } catch (e) {
    throwError(e);
  }
  return response;
};

const uploadFile = async (siteUrl, subsite, folder, file, formDigest, authorization) => {
  logger.progress('Uploading File'.padEnd(padGap), (`${folder}/${file.name}`).green);
  const folderPath = `('${folder}')/Files/add(overwrite=true, url='${file.name}')`;
  let response = null;
  try {
    response = await fetch(`${makePath(siteUrl, subsite)}/_api/Web/GetFolderByServerRelativeUrl${folderPath}`, {
      method: 'POST',
      body: file.data,
      headers: getHeaders(authorization, {
        'Content-Length': file.length,
        'X-RequestDigest': formDigest,
      }, ['Content-Type']),
    }).catch(throwError);
    response = await response.json().catch(throwError);
    if (response.error) {
      throw response.error;
    }
  } catch (e) {
    throwError(e);
  }
  return response;
};

const uploadSite = async (siteUrl, subsite, localFolder, remoteFolder, formDigest, authorization) => {
  logger.await('Uploading Site'.padEnd(padGap), (`${makePath(siteUrl, subsite)}/${remoteFolder}`).green);
  const availableFolders = [];
  try {
    const files = await listLocalFiles(localFolder, '').catch(throwError);
    for (const file of files) {
      const buff = readFile(file);
      const object = {
        name: normalizeFilePathForUpload(file.substring(file.lastIndexOf(path.sep) + 1)),
        length: buff.length,
        data: buff,
      };
      const absFile = file.replace(localFolder + path.sep, '');
      const subdir = absFile.substring(0, absFile.lastIndexOf(path.sep));
      if (availableFolders.indexOf(subdir) === -1) {
        availableFolders.push(subdir);
        await makeFolder(siteUrl, subsite, subdir, remoteFolder, formDigest, authorization).catch(throwError);
      }
      await uploadFile(siteUrl, subsite, `${remoteFolder}${subdir ? '/' : ''}${subdir}`, object, formDigest, authorization)
        .catch(throwError);
    }
  } catch (e) {
    throwError(e);
  }
  logger.complete('Uploaded Site'.padEnd(padGap), (`${makePath(siteUrl, subsite)}/${remoteFolder}`.green));
};

const deploy = async ({
  appClientId, appClientSecret, siteUrl, subsite, remoteFolder, distFolder, accessToken,
}) => {
  logger.start('Operation Started'.bold, '😈');
  try {
    const authorization = accessToken || await getAccessToken(siteUrl, subsite, appClientId, appClientSecret).catch(throwError);
    const formDigest = await getFormDigestValue(siteUrl, subsite, authorization).catch(throwError);
    await deleteSite(siteUrl, subsite, remoteFolder, formDigest, authorization).catch(throwError);
    await uploadSite(siteUrl, subsite, distFolder, remoteFolder, formDigest, authorization).catch(throwError);
    logger.success('Operation Completed'.bold, '😈');
  } catch (e) {
    throwError(e);
  }
};

exports.deploy = deploy;
