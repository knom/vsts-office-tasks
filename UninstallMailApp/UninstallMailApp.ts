// Upload just build task to VSTS...
// Upload extension VSTS...
// tfx extension create --manifest-globs .\vss-extension.json
// Upload to https://marketplace.visualstudio.com/manage/


// toolrunner help
// https://github.com/Microsoft/vsts-task-lib/blob/be60205671545ebef47d3a40569519da9b4d34b0/node/docs/vsts-task-lib.md

import q = require('q');
import tl = require('vsts-task-lib/task');
import trm = require('vsts-task-lib/toolrunner');
import path = require('path');
import fs = require('fs');
import xml2js = require('xml2js');
import ews = require('./lib/ews-soap/exchangeClient');

async function run() : q.Promise<void> {
    try {
		var ewsConnectedServiceName = tl.getInput('ewsConnectedServiceName', true);
		
		let appManifestXmlPath = tl.getPathInput('appManifestXmlPath', true, true);
		// let appManifestXmlPath = "C:\\Work\\TFS\\VSTS Office Manifest Uploader\\InstallMailApp\\SampleManifest.xml";

		fs.readFile(appManifestXmlPath, 'utf8', function (err, data) {
			if (err) {
				throw err;
			}

			let parser = new xml2js.Parser(
            {
                "explicitArray": false,
                "explicitRoot": false,
                "attrkey": "@"
            });

			parser.parseString(data, (err, result) => {
				let appId = result["Id"];

				tl.debug("app id:" + appId);
				
				let serverUrl = tl.getEndpointUrl(ewsConnectedServiceName, true);
				
				var ewsAuth = tl.getEndpointAuthorization(ewsConnectedServiceName, true);
				
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
								throw err;
							}
							else {
								tl.debug("Success.");
								tl.setResult(tl.TaskResult.Succeeded, "Successfully installed the app");
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
