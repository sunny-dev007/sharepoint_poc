var express = require('express');
var router = express.Router();

// Listed to import Controllers of All Resourse's Endpoints to use the Business Logic.
const sharepointControl = require("../controllers/sharepoint.controllers");

module.exports = app => {

    router.get('/sharepoint-sites/', sharepointControl.getSharepointSites);

    router.post('/sharepoint-folders/',sharepointControl.getSharepointFolderLists);

    router.post('/sharepoint-traverse/', sharepointControl.sharepointFolderTraverse);
    
    app.use('/api/v1/middleware/workspace', router);

};