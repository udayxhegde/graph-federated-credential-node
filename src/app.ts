import GraphFederatedClient from "./graph/graphfederatedclient";

const express = require("express");
var logHelper = require("./utils/loghelper");

require('dotenv-safe').config();

var graphClient = new GraphFederatedClient();

var port = process.env.PORT || 3001;

var app = express();
//
// initialize the logger
//
logHelper.init(app);

app.use(express.json());
app.use(express.urlencoded({ extended: true }));


app.get("/:appid", function(req:any, res:any) {
    let id = req.params.appid;

    logHelper.logger.info("calling get fed cred for %s", id);

    graphClient.getFederatedCredential(id)
    .then(function(response:any) {
        logHelper.logger.info("get response %o", response); 
        res.json(response);
    })
    .catch(function(error) {
        logHelper.logger.error("get error %o", error);
        res.status(500).send(error);        
    });
});


app.post("/:appid", function(req:any, res:any) {
    let id = req.params.appid;

    graphClient.setFederatedCredential(id, req.body.name, req.body.issuer, req.body.subject)
    .then(function(response:any) {
        res.json(response);
    })
    .catch(function(error) {
        logHelper.logger.error("set error %o", error);
        res.status(500).send(error);        
    });
});


app.listen(port);
logHelper.logger.info("express now running on poprt %d", port);

