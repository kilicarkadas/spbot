
var adalConfig = {
    'clientId': '9b9c2d70-2728-4244-8e83-9e2ecb55957f', // The client Id retrieved from the Azure AD App
    'clientSecret': '/Sf3YLWTJsgDbdJtSEOWR/aYCp9uDDVW0QJLZ/XgagI=', // The client secret retrieved from the Azure AD App
    'authorityHostUrl': 'https://login.microsoftonline.com/', // The host URL for the Microsoft authorization server
    'tenant': 'e825892c-c6b0-4a78-b8c2-df7888e8c689', // The tenant Id or domain name (e.g mydomain.onmicrosoft.com)
    'redirectUri': 'http://localhost:3978/api/oauthcallback', // This URL will be used for the Azure AD Application to send the authorization code.   
    'resource': 'https://groundteam.sharepoint.com', // The resource endpoint we want to give access to (in this case, SharePoint Online)
}
// Node fetch is the server version of whatwg-fetch
var fetch = require('node-fetch');

exports.createSiteGroups = (accessToken,siteUrl,groupName,groupDesc,groupRoleId) => {
       
    var spGroup = {
            "__metadata": {
                "type": "SP.Group"
            },
            "Title": groupName,
            "Description": groupDesc,
        };

    var p = new Promise((resolve, reject) => {
        var endpointUrl = siteUrl + "/_api/Web/SiteGroups";
        fetch(endpointUrl,{
            method: "POST",
            body: JSON.stringify(spGroup),
            headers: {
                "Authorization": "Bearer " + accessToken,
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose"
            }
        }).then(function (res) {
            return res.json();
        }).then(function (json) {
            //console.log(json);
            json.roleId=groupRoleId;
            resolve(json);
        }).catch(function (err) {
            reject(err);
        });
    });
    return p;
}

exports.assignRoleToSiteGroup = (accessToken,groupId,roleId) => {

    var p = new Promise((resolve, reject) => {

        var endpointUrl = adalConfig.resource + "/_api/web/roleassignments/addroleassignment(principalid="+ groupId +", roledefid="+ roleId +")";
        
        fetch(endpointUrl,{
            method: "POST",
            headers: {
                "Authorization": "Bearer " + accessToken,
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose"
            }
        }).then(function (res) {
            return res.json();
        }).then(function (json) {
            //console.log(json);
            resolve(json);
        }).catch(function (err) {
            reject(err);
        });
    });
    return p;
}



exports.searchSites = (query, accessToken) => {

    var p = new Promise((resolve, reject) => {

        var endpointUrl = adalConfig.resource + "/_api/search/query?querytext='" + query + "'";

        fetch(endpointUrl, {
            method: 'GET',
            headers: {
                "Authorization": "Bearer " + accessToken,
                "Accept": "application/json;odata=verbose"
            }
        }).then(function (res) {
            return res.json();
        }).then(function (json) {
            //console.log(json);
            resolve(json);
        }).catch(function (err) {
            reject(err);
        });
    });

    return p;
}

exports.getSiteCollections=(accessToken)=>{
    
    var p = new Promise((resolve, reject) => {

        var endpointUrl = adalConfig.resource + "/_api/search/query?querytext='contentclass:sts_site'";

        fetch(endpointUrl, {
            method: 'GET',
            headers: {
                "Authorization": "Bearer " + accessToken,
                "Accept": "application/json;odata=verbose"
            }
        }).then(function (res) {
            return res.json();
        }).then(function (json) {
            //console.log(json);
            resolve(json);
        }).catch(function (err) {
            reject(err);
        });
    });

    return p;

                        
}

exports.addNewSite = (siteTitle, siteDescription,accessToken) => {

    var p = new Promise((resolve, reject) => {
        var siteUrl = siteTitle.replace(/\s/g, "");
        var endpointUrl = adalConfig.resource + "/_api/web/webs/add";

        fetch(endpointUrl, {
            method: 'POST',
            body: JSON.stringify({
                'parameters': {
                    '__metadata': { 'type': 'SP.WebCreationInformation' },
                    'Url': siteUrl,
                    'Title': siteTitle,
                    'Description': siteDescription,
                    'WebTemplate': 'STS',
                    'UseSamePermissionsAsParentSite': true
                }
            }),
            headers: {
                "Authorization": "Bearer " + accessToken,
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose"
            }
        }).then(function (res) {
            return res.json();
        }).then(function (json) {
            //console.log(json);
            resolve(json);
        }).catch(function (err) {
            reject(err);
        });
    });

    return p;
}

exports.getSiteDetails=(siteTitle,accessToken)=>{
    var p = new Promise((resolve, reject) => {

        var endpointUrl = adalConfig.resource + "/" + siteTitle + '/_api/web/RoleAssignments/Groups';
        fetch(endpointUrl, {
            method: 'GET',
            headers: {
                "Authorization": "Bearer " + accessToken,
                "Accept": "application/json;odata=verbose"
            }
        }).then(function (res) {
            return res.json();
        }).then(function (json) {
            //console.log(json);
            resolve(json);
        }).catch(function (err) {
            reject(err);
        });
    });

    return p;

}