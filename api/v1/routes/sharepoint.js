const express = require('express');
const router = express.Router();

// Import SharePoint controller to handle the business logic for endpoints
const sharepointControl = require("../controllers/sharepoint.controllers");

module.exports = (app) => {
    // Endpoint to fetch SharePoint sites based on the domain name
    router.get('/sharepoint-sites/:sp_domain_name', sharepointControl.getSharepointSites);

    // Endpoint to fetch SharePoint folder lists
    router.post('/sharepoint-folders/', sharepointControl.getSharepointFolderLists);

    // Endpoint to traverse SharePoint folder structure
    router.post('/sharepoint-traverse/', sharepointControl.sharepointFolderTraverse);

    // Use the routes under the `/api/v1` prefix
    app.use('/api/v1', router);
};
