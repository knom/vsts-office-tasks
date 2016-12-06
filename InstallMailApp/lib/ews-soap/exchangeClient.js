"use strict";
var path = require("path");
var xml2js = require("xml2js");
var soap = require("soap");
var enumerable = require("linq");
var tl = require('vsts-task-lib/task');
var EWSClient = (function () {
    function EWSClient() {
        this.client = null;
    }
    EWSClient.prototype.initialize = function (settings, callback) {
        var _this = this;
        var endpoint = settings.url + "/EWS/Exchange.asmx";
        var url = path.join(__dirname, "Services.wsdl");
        var options = {
            escapeXML: false
        };
        soap.createClient(url, options, function (err, client) {
            if (err) {
                return callback(err);
            }
            if (!client) {
                return callback(new Error("Could not create client"));
            }
            _this.client = client;
            _this.client.addListener('request', function (xml) {
                tl.debug("---REQUEST---");
                tl.debug(xml);
                tl.debug("---END of REQUEST---");
            });
            _this.client.addListener('response', function (a, b) {
                tl.debug("---RESPONSE---");
                tl.debug("body: " + a);
                tl.debug("response: " + JSON.stringify(b));
                tl.debug("---END of RESPONSE---");
            });
            _this.client.addListener('soapError', function () {
                tl.warning("---SOAP ERROR---");
                tl.warning("---END of SOAP ERROR---");
            });
            if (settings.token) {
                _this.client.setSecurity(new soap.BearerSecurity(settings.token));
            }
            else {
                _this.client.setSecurity(new soap.BasicAuthSecurity(settings.username, settings.password));
            }
            return callback(null);
        }, endpoint);
    };
    EWSClient.prototype.installApp = function (manifest, callback) {
        if (!this.client) {
            return callback(new Error("Call initialize()"));
        }
        var soapRequest = 
        //"<tns:InstallApp xmlns:tns='http://schemas.microsoft.com/exchange/services/2006/messages'>" +
        "<Manifest>" + manifest + "</Manifest>"; //+
        //"</tns:InstallApp>";
        this.client.InstallApp(soapRequest, function (httpError, result, rawBody) {
            if (httpError) {
                if (httpError.response.statusCode && (httpError.response.statusCode == 401 || httpError.response.statusCode == 403)) {
                    return callback(new Error(httpError.response.statusCode + ": Unauthorized!"));
                }
                return callback(new Error(httpError));
            }
            var parser = new xml2js.Parser({
                "explicitArray": false,
                "explicitRoot": false,
                "attrkey": "@"
            });
            parser.parseString(rawBody, function (err, result) {
                var responseCode = result["s:Body"]["InstallAppResponse"]["ResponseCode"];
                if (responseCode !== "NoError") {
                    try {
                        var message = enumerable.from(result["s:Body"]["InstallAppResponse"]["MessageXml"]["t:Value"])
                            .where(function (i) { return i["@"]["Name"] === "InnerErrorMessageText"; })
                            .select(function (i) { return i["_"]; }).first();
                    }
                    catch (e) {
                        var message = "";
                    }
                    finally {
                        return callback(new Error(responseCode + ": " + JSON.stringify(message)));
                    }
                }
                callback(null);
            });
        });
    };
    ;
    EWSClient.prototype.uninstallApp = function (id, callback) {
        if (!this.client) {
            return callback(new Error("Call initialize()"));
        }
        var soapRequest = 
        //"<m:UninstallApp xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages'>" +
        "<ID>" + id + "</ID>"; // +
        //"</m:UninstallApp>";
        this.client.UninstallApp(soapRequest, function (httpError, result, rawBody) {
            if (httpError) {
                if (httpError.response.statusCode && (httpError.response.statusCode == 401 || httpError.response.statusCode == 403)) {
                    return callback(new Error(httpError.response.statusCode + ": Unauthorized!"));
                }
                return callback(new Error(httpError));
            }
            var parser = new xml2js.Parser({
                "explicitArray": false,
                "explicitRoot": false,
                "attrkey": "@"
            });
            parser.parseString(rawBody, function (err, result) {
                var responseCode = result["s:Body"]["UninstallAppResponse"]["ResponseCode"];
                if (responseCode !== "NoError") {
                    try {
                        var message = enumerable.from(result["s:Body"]["UninstallAppResponse"]["MessageXml"]["t:Value"])
                            .where(function (i) { return i["@"]["Name"] === "InnerErrorMessageText"; })
                            .select(function (i) { return i["_"]; }).first();
                    }
                    catch (e) {
                        var message = "";
                    }
                    finally {
                        return callback(new Error(responseCode + ": " + JSON.stringify(message)));
                    }
                }
                callback(null);
            });
        });
    };
    ;
    return EWSClient;
}());
exports.EWSClient = EWSClient;
