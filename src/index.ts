import { IOptions, IWebAndList, IFileInfo } from './utils/IDeployAppPkg';
import * as spauth from 'node-sp-auth';
import * as sprequest from 'sp-request';
import * as fs from 'fs';
import * as url from 'url';
import uuid4 from './helper/uuid4';

class DeployAppPkg {
  private _internalOptions: IOptions = {};
  
  constructor(options: IOptions) {
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
  
  public async start() {
    try {
      // Create the site URL
      const siteUrl = this._internalOptions.absoluteUrl ? this._internalOptions.absoluteUrl : `https://${this._internalOptions.tenant}.sharepoint.com/${this._internalOptions.site}`;
      
      // Specify the site credentials
      const credentials = {
        username: this._internalOptions.username,
        password: this._internalOptions.password
      };
      
      // Authenticate against SharePoint
      let request = sprequest.create(credentials);
      
      // Retrieve the site ID
      const siteId = await this._getSiteId(siteUrl, request);
      // Retrieve the web ID
      const webAndListInfo = await this._getWebAndListId(siteUrl, request);
      const webId = webAndListInfo.webId;
      const listId = webAndListInfo.listId;
      
      // Get the file information
      const fileInfo = await this._getFileInfo(siteUrl, request);
      
      // Retrieve the request-body.xml file
      let xmlReqBody = fs.readFileSync(__dirname + '/../request-body.xml', 'utf8');
      
      if(this._internalOptions.sp2016) {
        xmlReqBody = fs.readFileSync(__dirname + '/../request-body-SP2016.xml', 'utf8');
      }
      
      // Map all the required values to the XML body
      xmlReqBody = this._setXMLMapping(xmlReqBody, siteId, webId, listId, fileInfo, this._internalOptions.skipFeatureDeployment);
      // Post the request body to the processQuery endpoint
      await this._deployAppPkg(siteUrl, request, xmlReqBody, this._internalOptions.sp2016);
      
      if (this._internalOptions.verbose) {
        console.log('INFO: COMPLETED');
      }
      
      // Return the promise
      return;
    } catch (e) {
      console.log('ERROR:', e);
      throw "Something failed during the deployment";
    }
  }
  
  /**
  * Retrieve the site ID for the current URL
  * @param siteUrl The current site URL to call
  * @param request The authenticated sp-request function
  */
  private async _getSiteId(siteUrl: string, request: sprequest.ISPRequest) {
    return new Promise<string>((resolve, reject) => {
      const apiUrl = `${siteUrl}/_api/site?$select=Id`;
      
      return this._getRequest(apiUrl, request).then(result => {
        if (typeof result.Id !== "undefined" && result.id !== null) {
          if (this._internalOptions.verbose) {
            console.log(`INFO: Site ID - ${result.Id}`);
          }
          resolve(result.Id);
        } else {
          if (this._internalOptions.verbose) {
            console.log(`ERROR: ${JSON.stringify(result)}`);
          }
          reject('The site ID could not be retrieved');
        }
      });
    });
  }
  
  /**
  * Retrieve the web ID for the current URL
  * @param siteUrl The current site URL to call
  * @param authData The authenticated sp-request function
  */
  private async _getWebAndListId(siteUrl: string, request: sprequest.ISPRequest): Promise<IWebAndList> {
    return new Promise<IWebAndList>((resolve, reject) => {
      // Retrieve the relative site URL
      const relativeUrl: string = this._internalOptions.site === "" ? this._retrieveRelativeSiteUrl(siteUrl) : `/${this._internalOptions.site}`;
      // Create the API URL to call
      const apiUrl = `${siteUrl}/_api/web/getList('${relativeUrl}/appcatalog')?$select=Id,ParentWeb/Id&$expand=ParentWeb`;
      
      return this._getRequest(apiUrl, request).then(result => {
        if (typeof result.Id !== "undefined" && result.id !== null &&
        typeof result.ParentWeb !== "undefined" && result.ParentWeb !== null &&
        typeof result.ParentWeb.Id !== "undefined" && result.ParentWeb.Id !== null) {
          if (this._internalOptions.verbose) {
            console.log(`INFO: Web ID - ${result.ParentWeb.Id} / List ID - ${result.Id}`);
          }
          resolve({
            webId: result.ParentWeb.Id,
            listId: result.Id
          });
        } else {
          if (this._internalOptions.verbose) {
            console.log(`ERROR: ${JSON.stringify(result)}`);
          }
          reject('The web ID and list ID could not be retrieved');
        }
      });
    });
  }
  
  /**
  * Retrieve the file hidden version number and ID
  * @param siteUrl The current site URL to call
  * @param request The authenticated sp-request function
  */
  private async _getFileInfo(siteUrl: string, request: sprequest.ISPRequest): Promise<IFileInfo> {
    return new Promise<IFileInfo>((resolve, reject) => {
      const apiUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('AppCatalog')/Files('${this._internalOptions.filename}')?$expand=ListItemAllFields&$select=ListItemAllFields/Id,ListItemAllFields/owshiddenversion`;
      return this._getRequest(apiUrl, request).then(result => {
        if (typeof result.ListItemAllFields !== "undefined" && result.ListItemAllFields !== null &&
        typeof result.ListItemAllFields.Id !== "undefined" && result.ListItemAllFields.Id !== null &&
        typeof result.ListItemAllFields.owshiddenversion !== "undefined" && result.ListItemAllFields.owshiddenversion !== null) {
          if (this._internalOptions.verbose) {
            console.log(`INFO: List item ID - ${result.ListItemAllFields.Id} / version - ${result.ListItemAllFields.owshiddenversion}`);
          }
          resolve({
            id: result.ListItemAllFields.Id,
            version: result.ListItemAllFields.owshiddenversion
          });
        } else {
          if (this._internalOptions.verbose) {
            console.log(`ERROR: ${JSON.stringify(result)}`);
          }
          reject('The file information could not be retrieved');
        }
      });
    });
  }
  
  /**
  * Retrieve the file hidden version number and ID
  * @param siteUrl The current site URL to call
  * @param request The authenticated sp-request function
  */
  private async _getRequest(apiUrl: string, request: sprequest.ISPRequest): Promise<any> {           
    return new Promise<string>((resolve, reject) => {
      request.get(apiUrl)
      .then(response => {
        resolve(response.body.d);
      })
      .catch(err =>{
        if (this._internalOptions.verbose) {
          console.log('ERROR:', err);
        }
        reject(`Failed to call the API URL: ${apiUrl}`);
      });
    });
  }
  
  /**
  * Method to set the right mappings in the XML request body
  * @param xmlBody Contents of the XML file
  * @param siteId Site ID string 
  * @param webId Web ID string
  * @param listId List ID string
  * @param fileInfo File info: version number and the item ID 
  * @param skipDeployment Skip feature deployment
  */
  private _setXMLMapping(xmlBody: string, siteId: string, webId: string, listId: string, fileInfo: IFileInfo, skipDeployment: boolean): string {
    if (xmlBody) {
      // Replace the random token with a random guid
      xmlBody = xmlBody.replace(new RegExp('\\{randomId\\}', 'g'), uuid4.generate());
      // Replace the site ID token with the actual site ID string
      xmlBody = xmlBody.replace(new RegExp('\\{siteId\\}', 'g'), siteId);
      // Replace the web ID token with the actual web ID string
      xmlBody = xmlBody.replace(new RegExp('\\{webId\\}', 'g'), webId);
      // Replace the list ID token with the actual list ID string
      xmlBody = xmlBody.replace(new RegExp('\\{listId\\}', 'g'), listId);
      // Replace the item ID token with the actual item ID number
      xmlBody = xmlBody.replace(new RegExp('\\{itemId\\}', 'g'), fileInfo.id.toString());
      // Replace the file version token with the actual file version number
      xmlBody = xmlBody.replace(new RegExp('\\{fileVersion\\}', 'g'), fileInfo.version.toString());
      // Replace the skipFeatureDeployment token with the skipFeatureDeployment option
      xmlBody = xmlBody.replace(new RegExp('\\{skipFeatureDeployment\\}', 'g'), skipDeployment.toString());
      return xmlBody;
    } else {
      if (this._internalOptions.verbose) {
        console.log('ERROR:', xmlBody);
      }
      throw "Something wrong with the xmlBody";
    }
  }
  
  /**
  * Deploy the app package file
  * @param siteUrl The URL of the app catalog site
  * @param request The authenticated sp-request function
  */
  private async _deployAppPkg(siteUrl: string, request:sprequest.ISPRequest, xmlReqBody: string, sp2016: boolean) {
    return new Promise((resolve, reject) => {
      const apiUrl = `${siteUrl}/_vti_bin/client.svc/ProcessQuery`;
      
      request.requestDigest(siteUrl)
      .then(digest => {
        return request.post(apiUrl, {
          headers: {
            'X-RequestDigest': digest,
            'Content-Type': "application/xml"
          },
          body: xmlReqBody,
          json: false
        });
      })
      .then(response => {
        let body = JSON.parse(response.body);
        // Check if the current version of the app package is deployed
        if (sp2016 && body && body[2].IsClientSideSolutionDeployed) {
          if (this._internalOptions.verbose) {
            console.log('INFO: App package has been deployed to SP2016');
          }
          resolve(true);
        } else if (!sp2016 && body && body[2].IsClientSideSolutionCurrentVersionDeployed) {
          if (this._internalOptions.verbose) {
            console.log('INFO: App package has been deployed');
          }
          resolve(true);
        } else {
          if (this._internalOptions.verbose) {
            console.log('ERROR:', body);
          }
          reject('Failed to deploy the app package file.');
        }
      })
      .catch(err => {
        if (this._internalOptions.verbose) {
          console.log('ERROR:', err);
        }
        reject('Failed to deploy the app package file.');
        return;
      });
    });
  }
  
  /**
  * Retrieve the relative site path
  * @param siteUrl Absolute URL of the site
  */
  private _retrieveRelativeSiteUrl(siteUrl: string): string {
    const parsedUrl = url.parse(siteUrl);
    return parsedUrl.path;
  }
}

export const deploy = async (options: IOptions) => {
  return await new DeployAppPkg(options).start();
};