'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const spsync = require('gulp-spsync-creds').sync;
const sppkgDeploy = require('node-sppkg-deploy');
var gutil = require('gulp-util');
	
const environmentInfo = {
    "username": "admin@dev365x114392.onmicrosoft.com",
    "password": "",
    "tenant": "dev365x114392",
    "cdnSite": "sites/cdn",
    "cdnLib": "cdn/Sample",
    "catalogSite": "sites/appcatalog"
  };

build.task('upload-to-sharepoint', { 
   execute: (config) => {    
    environmentInfo.username = config.args['username'] || environmentInfo.username;
    environmentInfo.password = config.args['password'] || environmentInfo.password;
    environmentInfo.tenant = config.args['tenant'] || environmentInfo.tenant;
    environmentInfo.cdnSite = config.args['cdnsite'] || environmentInfo.cdnSite;
    environmentInfo.cdnLib = config.args['cdnlib'] || environmentInfo.cdnLib;
    
    gutil.log("Service user: " + environmentInfo.username);
    gutil.log("Trying to upload bundle to: " + `https://${environmentInfo.tenant}.sharepoint.com/${environmentInfo.cdnSite}`);

    return new Promise((resolve, reject) => {
        const deployFolder = require('./config/copy-assets.json');
        const folderLocation = `./${deployFolder.deployCdnPath}/**/*.*`;
        
        return gulp.src(folderLocation)
        .pipe(spsync({
            "username": environmentInfo.username,
            "password": environmentInfo.password,
            "site": `https://${environmentInfo.tenant}.sharepoint.com/${environmentInfo.cdnSite}`,
            "libraryPath": environmentInfo.cdnLib,
            "publish": true
        }))
        .on('finish', resolve);
    });
   }
});

build.task('upload-app-package', {
    execute: (config) => {      
        environmentInfo.username = config.args['username'] || environmentInfo.username;
        environmentInfo.password = config.args['password'] || environmentInfo.password;
        environmentInfo.tenant = config.args['tenant'] || environmentInfo.tenant;
        environmentInfo.catalogSite = config.args['catalogsite'] || environmentInfo.catalogSite;
     
        gutil.log("Service user: " + environmentInfo.username);
        gutil.log("Trying to upload package to: " + `https://${environmentInfo.tenant}.sharepoint.com/${environmentInfo.catalogSite}`);

        return new Promise((resolve, reject) => {
          const pkgFile = require('./config/package-solution.json');
          const folderLocation = `./sharepoint/${pkgFile.paths.zippedPackage}`;
     
          return gulp.src(folderLocation)
            .pipe(spsync({
              "username": environmentInfo.username,
              "password": environmentInfo.password,
              "site": `https://${environmentInfo.tenant}.sharepoint.com/${environmentInfo.catalogSite}`,
              "libraryPath": "AppCatalog",
              "publish": true
            }))
            .on('finish', resolve);
        });
    }
});

build.task('deploy-app-package', {
    execute: (config) => {
        environmentInfo.username = config.args['username'] || environmentInfo.username;
        environmentInfo.password = config.args['password'] || environmentInfo.password;
        environmentInfo.tenant = config.args['tenant'] || environmentInfo.tenant;
        environmentInfo.catalogSite = config.args['catalogsite'] || environmentInfo.catalogSite;
        
        gutil.log("Service user: " + environmentInfo.username);
        gutil.log("Trying to upload package to: " + `https://${environmentInfo.tenant}.sharepoint.com/${environmentInfo.catalogSite}`);

        const pkgFile = require('./config/package-solution.json');
        if (pkgFile) {
          let filename = pkgFile.paths.zippedPackage;
          filename = filename.split('/').pop();
          const skipFeatureDeployment = pkgFile.solution.skipFeatureDeployment ? pkgFile.solution.skipFeatureDeployment : false;
          return sppkgDeploy.deploy({
            username: environmentInfo.username,
            password: environmentInfo.password,
            tenant: environmentInfo.tenant,
            site: environmentInfo.catalogSite,
            filename: filename,
            skipFeatureDeployment: skipFeatureDeployment,
            verbose: true
          });
        }
    }
});

build.initialize(gulp);
