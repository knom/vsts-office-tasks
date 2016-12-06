// Upload just build task to VSTS...
// Upload extension VSTS...
// tfx extension create --manifest-globs .\vss-extension.json
// Upload to https://marketplace.visualstudio.com/manage/
"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator.throw(value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments)).next());
    });
};
// tsc InstallApplication.ts -w
// $env:INPUT_ewsConnectedServiceName = 'EP1'
// $env:ENDPOINT_URL_EP1 = 'https://mail.office365.com/'
// $env:ENDPOINT_AUTH_EP1 = '{ "parameters": { "username": "Some user", "password": "Some password" }, "scheme": "Some scheme" }'
// $env:ENDPOINT_DATA_EP1 = '{ "Key1": "Value1", "Key2", "Value2" }'
// 
// $env:INPUT_appManifestXmlPath="C:\...\a.xml"
// toolrunner help
// https://github.com/Microsoft/vsts-task-lib/blob/be60205671545ebef47d3a40569519da9b4d34b0/node/docs/vsts-task-lib.md
var q = require('q');
var tl = require('vsts-task-lib/task');
var fs = require('fs');
var ews = require('./lib/ews-soap/exchangeClient');
function run() {
    return __awaiter(this, void 0, q.Promise, function* () {
        try {
            var ewsConnectedServiceName = tl.getInput('ewsConnectedServiceName');
            var appManifestXmlPath = tl.getPathInput('appManifestXmlPath', true, true);
            fs.readFile(appManifestXmlPath, 'utf8', function (err, data) {
                if (err) {
                    throw err;
                }
                var request = require('request');
                require('request-debug')(request, function (type, data, request) {
                    tl.debug("---REQUEST-DEBUG " + type + "---");
                    tl.debug(JSON.stringify(data));
                    tl.debug("---END of REQUEST-DEBUG " + type + "---");
                });
                var manifest = new Buffer(data).toString("base64");
                tl.debug("manifest (base64): " + manifest);
                var serverUrl = tl.getEndpointUrl(ewsConnectedServiceName, true);
                var ewsAuth = tl.getEndpointAuthorization(ewsConnectedServiceName, true);
                var userName = ewsAuth['parameters']['username'];
                var password = ewsAuth['parameters']['password'];
                tl.debug("Initializing EWS client");
                var client = new ews.EWSClient();
                client.initialize({ url: serverUrl, username: userName, password: password }, function (err) {
                    if (err) {
                        throw err;
                    }
                    tl.debug("Calling InstallApp SOAP request");
                    client.installApp(manifest, function (err) {
                        if (err) {
                            tl.setResult(tl.TaskResult.Failed, err);
                        }
                        else {
                            tl.debug("Success.");
                            tl.setResult(tl.TaskResult.Succeeded, "Successfully installed the app");
                        }
                    });
                });
            });
        }
        catch (err) {
            // handle failures in one place
            tl.setResult(tl.TaskResult.Failed, err.message);
        }
    });
}
run();
