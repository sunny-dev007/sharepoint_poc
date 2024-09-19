const SharepointClass = require("../middleware/sharepoint.class");
const { v4: uuidv4 } = require("uuid");

// Instantiate SharePoint service class to use its methods
const sharepointCLS = new SharepointClass();

// Fetch SharePoint sites for the specified domain
exports.getSharepointSites = async (req, res) => {
  console.log("Fetching SharePoint sites...");

  try {
    const headers = {
      Authorization: `Bearer ${req.headers.sp_access_token}`,
      "Content-Type": "application/json",
    };
    const sp_domain_name = req.params.sp_domain_name;

    // Retrieve SharePoint sites using the service class
    const spDetails = await sharepointCLS.getSharepointSites(headers, sp_domain_name);

    if (spDetails.status) {
      return res.status(200).send({
        success: true,
        spDetails,
        message: "SharePoint sites fetched successfully.",
      });
    } else {
      return res.status(200).send({
        success: false,
        message: "Error retrieving SharePoint sites.",
      });
    }
  } catch (err) {
    return res.status(500).send({
      success: false,
      message: "Internal Server Error.",
      result: err,
    });
  }
};

// Fetch SharePoint folder lists for the specified site or item
exports.getSharepointFolderLists = async (req, res) => {
  try {
    const headers = {
      Authorization: `Bearer ${req.headers.sp_access_token}`,
      "Content-Type": "application/json",
    };
    const { site_id, item_id } = req.body;

    if (!item_id) {
      // Retrieve the drive ID for the given site
      const getDriveID = await sharepointCLS.getSharepointDriveID(site_id, headers);

      if (getDriveID.status) {
        // Fetch folder list using the drive ID
        const spDetails = await sharepointCLS.getSharepointFolderList(site_id, getDriveID.drive_id, headers);

        if (spDetails.status) {
          return res.status(200).send({
            success: true,
            spDetails,
            message: "Folder list fetched successfully.",
          });
        } else {
          return res.status(200).send({
            success: false,
            message: "Unauthorized access to SharePoint.",
          });
        }
      } else {
        return res.status(getDriveID.statusCode).send({
          status: false,
          statusCode: getDriveID.statusCode,
          message: getDriveID.data.error.message,
        });
      }
    } else {
      // Traverse the folder structure if an item ID is provided
      const spDetails = await sharepointCLS.sharepointFolderTraversing(site_id, item_id, headers);

      if (spDetails.status) {
        return res.status(200).send({
          success: true,
          spDetails,
          message: "Folder list fetched successfully.",
        });
      } else {
        return res.status(200).send({
          success: false,
          message: "Unauthorized access to SharePoint.",
        });
      }
    }
  } catch (err) {
    return res.status(500).send({
      success: false,
      message: "Internal Server Error.",
      result: err,
    });
  }
};

// Traverse SharePoint folder structure
exports.sharepointFolderTraverse = async (req, res) => {
  try {
    const headers = {
      Authorization: `Bearer ${req.headers.sp_access_token}`,
      "Content-Type": "application/json",
    };
    const { site_id, item_id } = req.body;

    // Traverse the folder structure using the SharePoint service
    const spFolderDetails = await sharepointCLS.sharepointFolderTraversing(site_id, item_id, headers);

    if (spFolderDetails.status) {
      return res.status(200).send({
        success: true,
        spFolderDetails,
        message: "Folder structure traversed successfully.",
      });
    } else {
      return res.status(200).send({
        success: false,
        message: "Unauthorized access to SharePoint.",
      });
    }
  } catch (err) {
    return res.status(500).send({
      success: false,
      message: "Internal Server Error.",
      result: err,
    });
  }
};