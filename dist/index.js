"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var _this = this;
Object.defineProperty(exports, "__esModule", { value: true });
var spauth = require("node-sp-auth");
var request = require("request");
var fs = require("fs");
var url = require("url");
var uuid4_1 = require("./helper/uuid4");
var DeployAppPkg = (function () {
    function DeployAppPkg(options) {
        this._internalOptions = {};
        this._internalOptions.username = options.username || "";
        this._internalOptions.password = options.password || "";
        this._internalOptions.tenant = options.tenant || "";
        this._internalOptions.site = options.site || "";
        this._internalOptions.absoluteUrl = options.absoluteUrl || "";
        this._internalOptions.filename = options.filename || "";
        this._internalOptions.sp2016 = options.sp2016 || false;
        this._internalOptions.skipFeatureDeployment = typeof options.skipFeatureDeployment !== "undefined" ? options.skipFeatureDeployment : true;
        this._internalOptions.verbose = typeof options.verbose !== "undefined" ? options.verbose : true;
        if (this._internalOptions.username === "") {
            throw "Username argument is required";
        }
        if (this._internalOptions.password === "") {
            throw "Password argument is required";
        }
        if (this._internalOptions.tenant === "" &&
            this._internalOptions.absoluteUrl === "") {
            throw "Tenant OR absoluteUrl argument is required";
        }
        if (this._internalOptions.site === "" &&
            this._internalOptions.absoluteUrl === "") {
            throw "Site OR absoluteUrl argument is required";
        }
        if (this._internalOptions.filename === "") {
            throw "Filename argument is required";
        }
    }
    DeployAppPkg.prototype.start = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                return [2, new Promise(function (resolve, reject) {
                        (function () { return __awaiter(_this, void 0, void 0, function () {
                            var siteUrl, credentials, authData, digestValue, siteId, webAndListInfo, webId, listId, fileInfo, xmlReqBody, e_1;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        _a.trys.push([0, 7, , 8]);
                                        siteUrl = this._internalOptions.absoluteUrl ? this._internalOptions.absoluteUrl : "https://" + this._internalOptions.tenant + ".sharepoint.com/" + this._internalOptions.site;
                                        credentials = {
                                            username: this._internalOptions.username,
                                            password: this._internalOptions.password
                                        };
                                        return [4, spauth.getAuth(siteUrl, credentials)];
                                    case 1:
                                        authData = _a.sent();
                                        authData.headers["Accept"] = "application/json";
                                        authData.headers["Content-type"] = "application/json";
                                        return [4, this._getDigestValue(siteUrl, authData)];
                                    case 2:
                                        digestValue = _a.sent();
                                        authData.headers["X-RequestDigest"] = digestValue;
                                        return [4, this._getSiteId(siteUrl, authData)];
                                    case 3:
                                        siteId = _a.sent();
                                        return [4, this._getWebAndListId(siteUrl, authData)];
                                    case 4:
                                        webAndListInfo = _a.sent();
                                        webId = webAndListInfo.webId;
                                        listId = webAndListInfo.listId;
                                        return [4, this._getFileInfo(siteUrl, authData)];
                                    case 5:
                                        fileInfo = _a.sent();
                                        xmlReqBody = fs.readFileSync(__dirname + '/../request-body.xml', 'utf8');
                                        if (this._internalOptions.sp2016) {
                                            xmlReqBody = fs.readFileSync(__dirname + '/../request-body-SP2016.xml', 'utf8');
                                        }
                                        xmlReqBody = this._setXMLMapping(xmlReqBody, siteId, webId, listId, fileInfo, this._internalOptions.skipFeatureDeployment);
                                        return [4, this._deployAppPkg(siteUrl, authData, xmlReqBody, this._internalOptions.sp2016)];
                                    case 6:
                                        _a.sent();
                                        if (this._internalOptions.verbose) {
                                            console.log('INFO: COMPLETED');
                                        }
                                        resolve();
                                        return [3, 8];
                                    case 7:
                                        e_1 = _a.sent();
                                        console.log('ERROR:', e_1);
                                        reject(e_1);
                                        return [3, 8];
                                    case 8: return [2];
                                }
                            });
                        }); })();
                    })];
            });
        });
    };
    DeployAppPkg.prototype._getDigestValue = function (siteUrl, authData) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                return [2, new Promise(function (resolve, reject) {
                        var apiUrl = siteUrl + "/_api/contextinfo?$select=FormDigestValue";
                        var requestOpts = authData.options;
                        requestOpts.headers = authData.headers;
                        requestOpts.url = apiUrl;
                        requestOpts.json = true;
                        request.post(requestOpts, function (err, resp, body) {
                            if (err) {
                                if (_this._internalOptions.verbose) {
                                    console.log('ERROR:', err);
                                }
                                reject('Failed to retrieve the site and web ID');
                                return;
                            }
                            if (body.FormDigestValue) {
                                if (_this._internalOptions.verbose) {
                                    console.log('INFO: FormDigestValue retrieved');
                                }
                                resolve(body.FormDigestValue);
                            }
                            else {
                                if (_this._internalOptions.verbose) {
                                    console.log('ERROR:', body);
                                }
                                reject('The FormDigestValue could not be retrieved');
                            }
                        });
                    })];
            });
        });
    };
    DeployAppPkg.prototype._getSiteId = function (siteUrl, authData) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                return [2, new Promise(function (resolve, reject) {
                        var apiUrl = siteUrl + "/_api/site?$select=Id";
                        return _this._getRequest(apiUrl, authData).then(function (result) {
                            if (typeof result.Id !== "undefined" && result.id !== null) {
                                if (_this._internalOptions.verbose) {
                                    console.log("INFO: Site ID - " + result.Id);
                                }
                                resolve(result.Id);
                            }
                            else {
                                if (_this._internalOptions.verbose) {
                                    console.log("ERROR: " + JSON.stringify(result));
                                }
                                reject('The site ID could not be retrieved');
                            }
                        });
                    })];
            });
        });
    };
    DeployAppPkg.prototype._getWebAndListId = function (siteUrl, authData) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                return [2, new Promise(function (resolve, reject) {
                        var relativeUrl = _this._internalOptions.site === "" ? _this._retrieveRelativeSiteUrl(siteUrl) : "/" + _this._internalOptions.site;
                        var apiUrl = siteUrl + "/_api/web/getList('" + relativeUrl + "/appcatalog')?$select=Id,ParentWeb/Id&$expand=ParentWeb";
                        return _this._getRequest(apiUrl, authData).then(function (result) {
                            if (typeof result.Id !== "undefined" && result.id !== null &&
                                typeof result.ParentWeb !== "undefined" && result.ParentWeb !== null &&
                                typeof result.ParentWeb.Id !== "undefined" && result.ParentWeb.Id !== null) {
                                if (_this._internalOptions.verbose) {
                                    console.log("INFO: Web ID - " + result.ParentWeb.Id + " / List ID - " + result.Id);
                                }
                                resolve({
                                    webId: result.ParentWeb.Id,
                                    listId: result.Id
                                });
                            }
                            else {
                                if (_this._internalOptions.verbose) {
                                    console.log("ERROR: " + JSON.stringify(result));
                                }
                                reject('The web ID and list ID could not be retrieved');
                            }
                        });
                    })];
            });
        });
    };
    DeployAppPkg.prototype._getFileInfo = function (siteUrl, headers) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                return [2, new Promise(function (resolve, reject) {
                        var apiUrl = siteUrl + "/_api/web/GetFolderByServerRelativeUrl('AppCatalog')/Files('" + _this._internalOptions.filename + "')?$expand=ListItemAllFields&$select=ListItemAllFields/Id,ListItemAllFields/owshiddenversion";
                        return _this._getRequest(apiUrl, headers).then(function (result) {
                            if (typeof result.ListItemAllFields !== "undefined" && result.ListItemAllFields !== null &&
                                typeof result.ListItemAllFields.Id !== "undefined" && result.ListItemAllFields.Id !== null &&
                                typeof result.ListItemAllFields.owshiddenversion !== "undefined" && result.ListItemAllFields.owshiddenversion !== null) {
                                if (_this._internalOptions.verbose) {
                                    console.log("INFO: List item ID - " + result.ListItemAllFields.Id + " / version - " + result.ListItemAllFields.owshiddenversion);
                                }
                                resolve({
                                    id: result.ListItemAllFields.Id,
                                    version: result.ListItemAllFields.owshiddenversion
                                });
                            }
                            else {
                                if (_this._internalOptions.verbose) {
                                    console.log("ERROR: " + JSON.stringify(result));
                                }
                                reject('The file information could not be retrieved');
                            }
                        });
                    })];
            });
        });
    };
    DeployAppPkg.prototype._getRequest = function (apiUrl, authData) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                return [2, new Promise(function (resolve, reject) {
                        var requestOpts = authData.options;
                        requestOpts.headers = authData.headers;
                        requestOpts.url = apiUrl;
                        requestOpts.json = true;
                        request(requestOpts, function (err, resp, body) {
                            if (err) {
                                if (_this._internalOptions.verbose) {
                                    console.log('ERROR:', err);
                                }
                                reject("Failed to call the API URL: " + apiUrl);
                                return;
                            }
                            resolve(body);
                        });
                    })];
            });
        });
    };
    DeployAppPkg.prototype._setXMLMapping = function (xmlBody, siteId, webId, listId, fileInfo, skipDeployment) {
        if (xmlBody) {
            xmlBody = xmlBody.replace(new RegExp('\\{randomId\\}', 'g'), uuid4_1.default.generate());
            xmlBody = xmlBody.replace(new RegExp('\\{siteId\\}', 'g'), siteId);
            xmlBody = xmlBody.replace(new RegExp('\\{webId\\}', 'g'), webId);
            xmlBody = xmlBody.replace(new RegExp('\\{listId\\}', 'g'), listId);
            xmlBody = xmlBody.replace(new RegExp('\\{itemId\\}', 'g'), fileInfo.id.toString());
            xmlBody = xmlBody.replace(new RegExp('\\{fileVersion\\}', 'g'), fileInfo.version.toString());
            xmlBody = xmlBody.replace(new RegExp('\\{skipFeatureDeployment\\}', 'g'), skipDeployment.toString());
            return xmlBody;
        }
        else {
            if (this._internalOptions.verbose) {
                console.log('ERROR:', xmlBody);
            }
            throw "Something wrong with the xmlBody";
        }
    };
    DeployAppPkg.prototype._deployAppPkg = function (siteUrl, authData, xmlReqBody, sp2016) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                return [2, new Promise(function (resolve, reject) {
                        var apiUrl = siteUrl + "/_vti_bin/client.svc/ProcessQuery";
                        authData.headers["Content-type"] = "application/xml";
                        var requestOpts = authData.options;
                        requestOpts.headers = authData.headers;
                        requestOpts.url = apiUrl;
                        requestOpts.json = true;
                        requestOpts.body = xmlReqBody;
                        request.post(requestOpts, function (err, resp, body) {
                            if (err) {
                                if (_this._internalOptions.verbose) {
                                    console.log('ERROR:', err);
                                }
                                reject('Failed to deploy the app package file.');
                                return;
                            }
                            if (sp2016 && body && body[2].IsClientSideSolutionDeployed) {
                                if (_this._internalOptions.verbose) {
                                    console.log('INFO: App package has been deployed to SP2016');
                                }
                                resolve(true);
                            }
                            else if (!sp2016 && body && body[2].IsClientSideSolutionCurrentVersionDeployed) {
                                if (_this._internalOptions.verbose) {
                                    console.log('INFO: App package has been deployed');
                                }
                                resolve(true);
                            }
                            else {
                                if (_this._internalOptions.verbose) {
                                    console.log('ERROR:', body);
                                }
                                reject('Failed to deploy the app package file.');
                            }
                        });
                    })];
            });
        });
    };
    DeployAppPkg.prototype._retrieveRelativeSiteUrl = function (siteUrl) {
        var parsedUrl = url.parse(siteUrl);
        return parsedUrl.path;
    };
    return DeployAppPkg;
}());
exports.deploy = function (options) { return __awaiter(_this, void 0, void 0, function () {
    var e_2;
    return __generator(this, function (_a) {
        switch (_a.label) {
            case 0:
                _a.trys.push([0, 2, , 3]);
                return [4, new DeployAppPkg(options).start()];
            case 1: return [2, _a.sent()];
            case 2:
                e_2 = _a.sent();
                return [3, 3];
            case 3: return [2];
        }
    });
}); };
//# sourceMappingURL=index.js.map