"use strict";
var path = require("path");
var xml2js = require("xml2js");
var soap = require("soap");
var EWSClient = (function () {
    function EWSClient() {
        this.client = null;
    }
    EWSClient.prototype.initialize = function (settings, callback) {
        var _this = this;
        var endpoint = settings.url + "/EWS/Exchange.asmx";
        var url = path.join(__dirname, "Services.wsdl");
        soap.createClient(url, {}, function (err, client) {
            if (err) {
                return callback(err);
            }
            if (!client) {
                return callback(new Error("Could not create client"));
            }
            _this.client = client;
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
        var soapRequest = "<tns:InstallApp xmlns:tns='http://schemas.microsoft.com/exchange/services/2006/messages'>" +
            "<tns:Manifest>" + manifest + "</tns:Manifest>" +
            "</tns:InstallApp>";
        this.client.InstallApp(soapRequest, function (err, result, body) {
            if (err) {
                if (result.statusCode && (result.statusCode == 401 || result.statusCode == 403)) {
                    return callback("Unauthorized!");
                }
                return callback(err);
            }
            var parser = new xml2js.Parser({
                "explicitArray": false,
                "explicitRoot": false,
                "attrkey": "@"
            });
            parser.parseString(body, function (err, result) {
                var responseCode = result["s:Body"]["InstallAppResponse"]["ResponseCode"];
                if (responseCode !== "NoError") {
                    return callback(new Error(responseCode));
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
        var soapRequest = "<m:UninstallApp xmlns:m='http://schemas.microsoft.com/exchange/services/2006/messages'>" +
            "<m:ID>" + id + "</m:ID>" +
            "</m:UninstallApp>";
        this.client.UninstallApp(soapRequest, function (err, result, body) {
            if (err) {
                if (result.statusCode && (result.statusCode == 401 || result.statusCode == 403)) {
                    return callback("Unauthorized!");
                }
                return callback(err);
            }
            var parser = new xml2js.Parser({
                "explicitArray": false,
                "explicitRoot": false,
                "attrkey": "@"
            });
            parser.parseString(body, function (err, result) {
                var responseCode = result["s:Body"]["UninstallAppResponse"]["ResponseCode"];
                if (responseCode !== "NoError") {
                    return callback(new Error(responseCode));
                }
                callback(null);
            });
        });
    };
    ;
    return EWSClient;
}());
exports.EWSClient = EWSClient;
