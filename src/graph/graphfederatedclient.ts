import {Client} from "@microsoft/microsoft-graph-client";
import {TokenCredentialAuthenticationProvider} from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import {ClientSecretCredential} from "@azure/identity";


const tokenScopes = {
    scopes: ["https://graph.microsoft.com/.default"]
};

var logger = require("../utils/loghelper").logger;
require('isomorphic-fetch');


class GraphFederatedClient {
    graphClient: Client;

    constructor() {
        let clientId:any = process.env.USE_CLIENT_ID;
        let tenantId:any = process.env.USE_TENANT_ID;
        let clientSecret:any = process.env.USE_SECRET;
        logger.info("client %s tenant %s", clientId, tenantId);
        const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        const authProvider = new TokenCredentialAuthenticationProvider(credential, tokenScopes);
        this.graphClient = Client.initWithMiddleware({ 
            debugLogging: true,
            defaultVersion: 'beta',
            authProvider
        });
    }

    async setFederatedCredential(appId:string, idName:string, issuer:string, subject:string) 
    {

        const newCredential = {
            name : idName,
            issuer : issuer,
            subject: subject,
            audience : "api://AzureADTokenExchange"
        };

        return this.graphClient.api(`/applications/${appId}/federatedIdentityCredentials`).post(newCredential)
        .then(function(response:any) {
            logger.info("setFederatedCredential response %o", response)
            return response;
        })
        .catch(function(error:any) {
            logger.info("setFederatedClient error %o", error);
            throw(error);

        });
    }

    async getFederatedCredential(appId:string, _id?:string)
    {
        return this.graphClient.api(`/applications/${appId}/federatedIdentityCredentials/`).get()
        .then(function(response:any) {
            logger.info("getFederatedCredential response %o", response)
            return response;
        })
        .catch(function(error:any) {
            logger.info("getFederatedClient error %o", error);
            throw(error);

        });
    }
}

export default GraphFederatedClient;
