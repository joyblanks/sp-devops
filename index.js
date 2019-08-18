#! /usr/bin/env node

require('dotenv').config();
const argv = require('minimist')(process.argv.slice(2));

require('./src/main').main({...process.env, ...argv});

// You can pass any of these items from command line as well
// $ sp-devops --deploy --sp-subsite=my-subsite
