const SharepointClass = require("../middleware/sharepoint.class");
const { v4: uuidv4 } = require("uuid");

// Create an Object of Resource Class use to use All methods.
const sharepointCLS = new SharepointClass();

exports.getSharepointSites = async (req, res) => {
  console.log("Fetching sharepoint sites....");

  try {
    // set variables
    const headers = {
      Authorization: `Bearer ${req.headers.sp_access_token}`,
      "Content-Type": "application/json",
    };

    // prep headers for azure
    const az_headers = {
      sp_access_token: req.headers.sp_access_token,
    };

    // check if sharepoint access
    const spDetails = await sharepointCLS.getSharepointSites(headers);

    if (spDetails.status == true) {
      res.status(200).send({
        success: true,
        spDetails,
        message: "Response has been fetched",
      });
    } else {
      res.status(200).send({
        success: false,
        message: "Error uploading file",
      });
    }
  } catch (err) {
    res.status(500).send({
      success: false,
      message: "Error uploading file",
      result: err,
    });
  }
};

exports.getSharepointFolderLists = async (req, res) => {

  try {
    // set variables
    const headers = {
      Authorization: `Bearer ${req.headers.sp_access_token}`,
      "Content-Type": "application/json",
    };

    // GET BODY VALUES
    const { workspace_id, site_id, item_id } = req.body;

    // prep headers for azure
    const az_headers = {
      sp_access_token: req.headers.sp_access_token,
    };

    if (item_id == "") {
      // console.log("Calling.... Folder List::: NOT Found Item ID ");
      // get drive id
      const getDriveID = await sharepointCLS.getSharepointDriveID(
        workspace_id,
        site_id,
        headers
      );
      // console.log('getDriveID, ',  getDriveID);
      if (getDriveID.status == true) {
        // check if sharepoint access
        const spDetails = await sharepointCLS.getSharepointFolderList(
          workspace_id,
          site_id,
          getDriveID.drive_id,
          user_id,
          headers
        );

        if (spDetails.status == true) {
          res.status(200).send({
            success: true,
            spDetails,
            message: "Response has been fetched",
          });
        } else {
          res.status(200).send({
            success: false,
            message: "Unauthorized access to Sharepoint",
          });
        }
      } else {
        res.status(getDriveID.statusCode).send({
          status: false,
          statusCode: getDriveID.statusCode,
          message: getDriveID.data.error.message,
        });
      }
    } else {
      // check if sharepoint access
      const spDetails = await sharepointCLS.sharepointFolderTraversing(
        workspace_id,
        site_id,
        item_id,
        user_id,
        headers
      );

      if (spDetails.status == true) {
        res.status(200).send({
          success: true,
          spDetails,
          message: "Response has been fetched",
        });
      } else {
        res.status(200).send({
          success: false,
          message: "Unauthorized access to Sharepoint",
        });
      }
    }
  } catch (err) {
    res.status(500).send({
      success: false,
      message: "Unauthorized access to Sharepoint",
      result: err,
    });
  }
};

exports.sharepointFolderTraverse = async (req, res) => {
  // console.log("Fetching sharepoint sites....");
  // get user
  const user_id = await getUserAccount(req);

  try {
    // set variables
    const headers = {
      Authorization: `Bearer ${req.headers.sp_access_token}`,
      "Content-Type": "application/json",
    };

    // GET BODY VALUES
    const { site_id, item_id } = req.body;

    // prep headers for azure
    const az_headers = {
      sp_access_token: req.headers.sp_access_token,
    };

    // check if sharepoint access
    const spFolderDetails = await sharepointCLS.sharepointFolderTraversing(
      site_id,
      item_id,
      user_id,
      headers
    );

    if (spFolderDetails.status == true) {
      res.status(200).send({
        success: true,
        spFolderDetails,
        message: "Response has been fetched",
      });
    } else {
      res.status(200).send({
        success: false,
        message: "Unauthorized access to Sharepoint",
      });
    }
  } catch (err) {
    res.status(500).send({
      success: false,
      message: "Unauthorized access to Sharepoint",
      result: err,
    });
  }
};

