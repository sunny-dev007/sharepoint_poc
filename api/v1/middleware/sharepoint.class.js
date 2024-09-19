const axios = require("axios");

// All Business logic will be here
class SharepointClass {
  constructor() {}

  // get sharepoint sites name details
  async getSharepointSites(headers) {
    // get site names
    const getSites = `https://graph.microsoft.com/v1.0/sites/hanoverresearch.sharepoint.com/sites?search=*&top=999`;
    try {
      const siteIDresp = await axios.get(getSites, { headers: headers });
      // check the @next link in future - to be done
      const site_objs = siteIDresp.data;
      let type = "site";

      let site_names = [];
      let childrenR = [];
      site_objs.value.map((ele) => {
        // console.log(ele);
        if ("root" in ele && ele.root && Object.keys(ele.root).length > 0) {
          childrenR = ele.root;
        }
        const pushEle = {
          createdDateTime: ele.createdDateTime,
          value: ele.id,
          lastModifiedDateTime: ele.lastModifiedDateTime,
          label: ele.displayName,
          type: type,
          children: childrenR,
        };
        site_names.push(pushEle);
      });

      // Grab the One Drive Site ID from the Graph API and Push it to the Sharepoint Sites list
      const getOneFriveObject = await this.getOneDriveSitesID(headers);
      site_names.push(getOneFriveObject.oneDriveObject)

      // To Sort the Labels
      site_names.sort((a,b) => (a.label > b.label) ? 1 : ((b.label > a.label) ? -1 : 0));

      if (site_names) {
        return {
          status: true,
          site_names,
        };
      } else {
        return { status: false };
      }
    } catch (err) {
      console.log(err);
      throw new Error("Error to access the Sharepoint boundries");
    }
  }

  // get sharepoint sites name details
  async getSharepointDriveID(workspace_id, site_id, headers) {
    // get site names
    const requestToDriveID = `https://graph.microsoft.com/v1.0/sites/${site_id}/drive?search=*&top=999`;
    try {
      const getDriveID = await axios.get(requestToDriveID, {
        headers: headers,
      });
      //console.log('getDriveID =>', getDriveID)
      if (getDriveID) {
        return {
          status: true,
          statusCode: getDriveID.status,
          drive_id:getDriveID.data.id
        };
      }
    } catch (err) {
      console.log('GetDrive Catch Error', err);
      return { 
        status: false,
        statusCode: err.response.status,
        message: err.response.statusText,
        data: err.response.data 
      };
    }
  }

  // get sharepoint sites name details
  async getSharepointFolderList(
    workspace_id,
    site_id,
    drive_id,
    user_id,
    headers
  ) {
    // get site names
    const getFolderLevels = `https://graph.microsoft.com/v1.0/sites/${site_id}/drives/${drive_id}/root/children?search=*&top=999`;
    try {
      const folderLvl = await axios.get(getFolderLevels, { headers: headers });
      const folders_objs = folderLvl.data;

      //console.log('folders_objs =>', folders_objs)

      let filtered_folders_objs = [];

      if (folders_objs.value && folders_objs.value.length > 0) {
        // Filter out folders with package type oneNote
        filtered_folders_objs = folders_objs.value.filter(folder => !(folder.package && folder.package.type === 'oneNote'));
      }

      //console.log('filtered_folders_objs =>', filtered_folders_objs)

      let folder_level = [];
      let type = "";
      let iterateSupportFile = false;
      let fileExists = false;
      let isFileProcessed = null;
      let isFileModified = false;
      let file_format = 'unknown';

      for (let i = 0; i < filtered_folders_objs.length; i++) {
        const ele = filtered_folders_objs[i];
        //console.log('First Level Folder Details OBJ==>>>>>', ele);

        const supportedFilesMime = [
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          "text/plain",
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          "application/pdf",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "text/csv",
          "application/vnd.ms-excel"
        ];

          const hasFile = Object.keys(ele).includes("file");
          const hasFolder = Object.keys(ele).includes("folder");

          // Login to filter unsupported files
          if (hasFolder == true) {
            type = "Folder";
            iterateSupportFile = true;
            const pushEle = {
              createdDateTime: ele.createdDateTime,
              parentNodeID: site_id,
              value: ele.id,
              lastModifiedDateTime: ele.lastModifiedDateTime,
              label: ele.name,
              webUrl: ele.webUrl,
              type: type,
              isFileModified: isFileModified,
              isFileProcessed: isFileProcessed,
              fileExists: fileExists,
              children: [],
            };
            folder_level.push(pushEle);
          } else if (hasFile == true) {
            if(supportedFilesMime.includes(ele.file.mimeType)){
              if(ele.file.mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation'){
                file_format = 'pptx';
              }else if(ele.file.mimeType === 'text/plain'){
                file_format = 'txt';
              }else if(ele.file.mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'){
                file_format = 'docx';
              }else if(ele.file.mimeType === 'application/pdf'){
                file_format = 'pdf';
              }else if(ele.file.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'){
                file_format = 'xlsx';
              }else if(ele.file.mimeType === 'text/csv'){
                file_format = 'csv';
              }else if(ele.file.mimeType === 'application/vnd.ms-excel') {
                file_format = 'csv';
              }
              const isFileExist = await this.isAvalFileInWorkspace(
                workspace_id,
                site_id,
                ele.id,
                ele.lastModifiedDateTime,
                user_id
              );
              //console.log("isFileExist ===>", isFileModified);
              isFileProcessed = isFileExist.isFileProcessed;
              if (isFileExist.isFileModified == true) {
                isFileModified = true;
              } else {
                isFileModified = false;
              }
  
              if (isFileExist.fileExists == true) {
                fileExists = true;
              } else {
                fileExists = false;
              }
  
              // if(isFileExist.isFileProcessed == true){
              //   isFileProcessed = true;
              // }else{
              //   isFileProcessed = false;
              // }
              //console.log("hasFile =>", hasFile);
              type = "File";
              const pushEle = {
                createdDateTime: ele.createdDateTime,
                parentNodeID: site_id,
                value: ele.id,
                lastModifiedDateTime: ele.lastModifiedDateTime,
                label: ele.name,
                webUrl: ele.webUrl,
                type: type,
                isFileModified: isFileModified,
                isFileProcessed: isFileProcessed,
                fileExists: fileExists,
                file_format:file_format,
                children: [],
              };
              folder_level.push(pushEle);
            }
          }
      }

      if (folder_level) {
        return {
          status: true,
          folder_level,
        };
      } else {
        return { status: false };
      }
    } catch (err) {
      console.log(err);
      throw new Error("Error in sharepoint access");
    }
  }

  // get sharepoint sites name details
  async sharepointFolderTraversing(
    workspace_id,
    site_id,
    item_id,
    user_id,
    headers
  ) {
    // get site names
    const getFolderLevels = `https://graph.microsoft.com/v1.0/sites/${site_id}/drive/items/${item_id}/children?search=*&top=999`;
    try {
      const folderTraverse = await axios.get(getFolderLevels, {
        headers: headers,
      });
      const folders_objs = folderTraverse.data;
      //console.log("Folder Traversing =>", folders_objs);

      let fileExists = false;
      let isFileProcessed = null;
      let isFileModified = false;
      let folder_level = [];
      let type = "";
      let iterateSupportFile = false;
      let file_format = 'unknown';

      for (let i = 0; i < folders_objs.value.length; i++) {
        const ele = folders_objs.value[i];
        //console.log("Folder Traversing =>", ele);
        //console.log("Folder Traversing =>", ele.file);

        const supportedFilesMime = [
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          "text/plain",
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          "application/pdf",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "text/csv",
          "application/vnd.ms-excel"
        ];

          // Pushes the Supported files object into the rows********************
          const hasFile = Object.keys(ele).includes("file");
          const hasFolder = Object.keys(ele).includes("folder");

          // Login to filter unsupported files
          if (hasFolder == true) {
            // console.log("hasFolder =>", hasFolder);
            type = "Folder";
            iterateSupportFile = true;

            const pushEle = {
              createdDateTime: ele.createdDateTime,
              parentNodeID: item_id,
              site_id: site_id,
              value: ele.id,
              lastModifiedDateTime: ele.lastModifiedDateTime,
              label: ele.name,
              webUrl: ele.webUrl,
              fileExists: fileExists,
              isFileModified: isFileModified,
              isFileProcessed: isFileProcessed,
              downloadURL: ele["@microsoft.graph.downloadUrl"],
              type: type,
              children: [],
            };
            folder_level.push(pushEle);

          } else if (hasFile == true) {
            if(supportedFilesMime.includes(ele.file.mimeType)){
              if(ele.file.mimeType === 'application/vnd.openxmlformats-officedocument.presentationml.presentation'){
                file_format = 'pptx';
              }else if(ele.file.mimeType === 'text/plain'){
                file_format = 'txt';
              }else if(ele.file.mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'){
                file_format = 'docx';
              }else if(ele.file.mimeType === 'application/pdf'){
                file_format = 'pdf';
              }else if(ele.file.mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'){
                file_format = 'xlsx';
              }else if(ele.file.mimeType === 'text/csv'){
                file_format = 'csv';
              }else if(ele.file.mimeType === 'application/vnd.ms-excel') {
                file_format = 'csv';
              }

              const isFileExist = await this.isAvalFileInWorkspace(
                workspace_id,
                site_id,
                ele.id,
                ele.lastModifiedDateTime,
                user_id
              );
              // console.log("isFileExist ===>", isFileModified);
              isFileProcessed = isFileExist.isFileProcessed;
              if (isFileExist.isFileModified == true) {
                isFileModified = true;
              } else {
                isFileModified = false;
              }

              if (isFileExist.fileExists == true) {
                fileExists = true;
              } else {
                fileExists = false;
              }

              // if(isFileExist.isFileProcessed == true){
              //   isFileProcessed = true;
              // }else{
              //   isFileProcessed = false;
              // }
              // console.log("hasFile =>", hasFile);
              type = "File";
              const pushEle = {
                createdDateTime: ele.createdDateTime,
                parentNodeID: item_id,
                site_id: site_id,
                value: ele.id,
                lastModifiedDateTime: ele.lastModifiedDateTime,
                label: ele.name,
                webUrl: ele.webUrl,
                fileExists: fileExists,
                isFileModified: isFileModified,
                isFileProcessed: isFileProcessed,
                file_format:file_format,
                downloadURL: ele["@microsoft.graph.downloadUrl"],
                type: type,
                children: [],
              };
              folder_level.push(pushEle);
            }
          }
      }

      if (folder_level) {
        return {
          status: true,
          folder_level,
        };
      } else {
        return { status: false };
      }
    } catch (err) {
      console.log(err);
      throw new Error("Error in sharepoint access");
    }
  }

}

module.exports = SharepointClass;
