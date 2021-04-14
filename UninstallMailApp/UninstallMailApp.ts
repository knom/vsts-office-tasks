// Upload just build task to VSTS...
// Upload extension VSTS...
// tfx extension create --manifest-globs .\vss-extension.json
// Upload to https://marketplace.visualstudio.com/manage/

// tsc InstallApplication.ts -w
// $env:INPUT_ewsConnectedServiceName = 'EP1'
// $env:ENDPOINT_URL_EP1 = 'https://mail.office365.com/'
// $env:ENDPOINT_AUTH_EP1 = '{ "parameters": { "username": "Some user", "password": "Some password" }, "scheme": "Some scheme" }'
// $env:ENDPOINT_DATA_EP1 = '{ "Key1": "Value1", "Key2", "Value2" }'
// 
// $env:INPUT_appManifestXmlPath="C:\...\a.xml"

// toolrunner help
// https://github.com/Microsoft/vsts-task-lib/blob/be60205671545ebef47d3a40569519da9b4d34b0/node/docs/vsts-task-lib.md

import q = require('q');
import tl = require('azure-pipelines-task-lib/task');
import trm = require('azure-pipelines-task-lib/toolrunner');
import path = require('path');
import fs = require('fs');
import xml2js = require('xml2js');
import ews = require('./lib/ews-soap/exchangeClient');

function run() {
    try {
		var ewsConnectedServiceName = tl.getInput('ewsConnectedServiceName', true);
		
		let appManifestXmlPath = tl.getPathInput('appManifestXmlPath', true, true);
		
		fs.readFile(appManifestXmlPath, 'utf8', function (err, data) {
			if (err) {
				throw err;
			}
			
			var request = require('request');
			require('request-debug')(request, function(type:any, data:any, request:any) {
				tl.debug("---REQUEST-DEBUG " + type + "---");
				tl.debug(JSON.stringify(data));
				tl.debug("---END of REQUEST-DEBUG " + type + "---");
			});

			let parser = new xml2js.Parser(
            {
                "explicitArray": false,
                "explicitRoot": false,
                "attrkey": "@"
            });

			parser.parseString(data, (err:any, result:any) => {
				let appId = result["Id"];

				tl.debug("app id:" + appId);
				
				let serverUrl = tl.getEndpointUrl(ewsConnectedServiceName, true);
				
				let ewsAuth = tl.getEndpointAuthorization(ewsConnectedServiceName, true);
				
				if (!serverUrl || !ewsAuth)
				{
					throw {message:"Endpoint is not set!"};
				}
				
				let userName = ewsAuth['parameters']['username'];
				let password = ewsAuth['parameters']['password'];
				
				tl.debug("Initializing EWS client");
				let client = new ews.EWSClient();
				client.initialize({ url: serverUrl, username: userName, password: password },
					(err: any) => {
						if (err){
							throw err;
						}

						tl.debug("Calling UninstallApp SOAP request");
						client.uninstallApp(appId, (err: any) => {
							if (err) {
								tl.setResult(tl.TaskResult.Failed, err);
							}
							else {
								tl.debug("Success.");
								tl.setResult(tl.TaskResult.Succeeded, "Successfully uninstalled the app");
							}
						});
					});
			});
    	});
	}
    catch (err) {
        // handle failures in one place
        tl.setResult(tl.TaskResult.Failed, err.message);
    }
}

run();
