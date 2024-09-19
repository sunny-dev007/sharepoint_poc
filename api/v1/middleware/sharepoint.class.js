const axios = require("axios");

// Class containing business logic for SharePoint operations
class SharepointClass {
  constructor() {}

  // Get list of SharePoint sites based on the domain
  async getSharepointSites(headers, sp_domain) {
    const getSitesUrl = `https://graph.microsoft.com/v1.0/sites/${sp_domain}/sites?search=*&top=999`;
    try {
      const response = await axios.get(getSitesUrl, { headers });
      const siteData = response.data;
      const siteList = siteData.value.map(site => ({
        createdDateTime: site.createdDateTime,
        site_id: site.id,
        lastModifiedDateTime: site.lastModifiedDateTime,
        label: site.displayName,
        type: "site",
      }));

      // Sort site names alphabetically
      siteList.sort((a, b) => a.label.localeCompare(b.label));

      return { status: true, site_names: siteList };
    } catch (err) {
      console.error("Error fetching SharePoint sites:", err);
      throw new Error("Error accessing SharePoint boundaries");
    }
  }

  // Get SharePoint Drive ID for the specified site
  async getSharepointDriveID(site_id, headers) {
    const getDriveUrl = `https://graph.microsoft.com/v1.0/sites/${site_id}/drive?search=*&top=999`;
    try {
      const response = await axios.get(getDriveUrl, { headers });
      return {
        status: true,
        statusCode: response.status,
        drive_id: response.data.id,
      };
    } catch (err) {
      console.error("Error fetching Drive ID:", err);
      return {
        status: false,
        statusCode: err.response.status,
        message: err.response.statusText,
        data: err.response.data,
      };
    }
  }

  // Get folder list within a SharePoint drive
  async getSharepointFolderList(site_id, drive_id, headers) {
    const getFoldersUrl = `https://graph.microsoft.com/v1.0/sites/${site_id}/drives/${drive_id}/root/children?search=*&top=999`;
    try {
      const response = await axios.get(getFoldersUrl, { headers });
      const folders = response.data.value.filter(folder => !(folder.package && folder.package.type === 'oneNote'));

      const folderList = folders.map(folder => ({
        createdDateTime: folder.createdDateTime,
        parentNodeID: site_id,
        item_id: folder.id,
        lastModifiedDateTime: folder.lastModifiedDateTime,
        label: folder.name,
        webUrl: folder.webUrl,
        type: folder.folder ? "Folder" : "File",
      }));

      return { status: true, folder_level: folderList };
    } catch (err) {
      console.error("Error fetching folder list:", err);
      throw new Error("Error accessing SharePoint");
    }
  }

  // Traverse a SharePoint folder's structure
  async sharepointFolderTraversing(site_id, item_id, headers) {
    const traverseUrl = `https://graph.microsoft.com/v1.0/sites/${site_id}/drive/items/${item_id}/children?search=*&top=999`;
    try {
      const response = await axios.get(traverseUrl, { headers });
      const items = response.data.value;

      const folderLevel = items.map(item => ({
        createdDateTime: item.createdDateTime,
        parentNodeID: item_id,
        site_id,
        item_id: item.id,
        lastModifiedDateTime: item.lastModifiedDateTime,
        label: item.name,
        webUrl: item.webUrl,
        type: item.folder ? "Folder" : "File",
        downloadURL: item["@microsoft.graph.downloadUrl"],
      }));

      return { status: true, folder_level: folderLevel };
    } catch (err) {
      console.error("Error traversing folder:", err);
      throw new Error("Error accessing SharePoint");
    }
  }
}

module.exports = SharepointClass;
