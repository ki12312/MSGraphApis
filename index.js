/**
 * Imports
 */
require("dotenv").config();
const express = require("express");
const app = express();
const axios = require("axios");
const cors = require("cors");
const authenticateToken = require("./src/middleware/authentication");
const jwt = require("jsonwebtoken");
const https = require("https");
const qs = require("qs");
const fs = require('fs');
const { jwtDecode } =require("jwt-decode");


/**
 * Get Access Token Directly
 */
const getAccessToken = async () => {
  // Make sure you replace these values from the copied values of your app
  const APP_ID = process.env.CLIENT_ID;
  console.log("App_ID = ", APP_ID);
  const APP_SECRET = process.env.CLIENT_SECRET;
  console.log("APP_SECRET = ",APP_SECRET)
  const TOKEN_ENDPOINT = `https://login.microsoftonline.com/44f47d0f-9269-44ca-9bbd-8d7df9cf7363/oauth2/v2.0/token`;
  const MS_GRAPH_SCOPE = "https://graph.microsoft.com/.default";

  const postData = {
    client_id: APP_ID,
    scope: MS_GRAPH_SCOPE,
    client_secret: APP_SECRET,
    grant_type: "client_credentials",
  };

  axios.defaults.headers.post["Content-Type"] =
    "application/x-www-form-urlencoded";

  return axios
    .post(TOKEN_ENDPOINT, qs.stringify(postData))
    .then((response) => {

      return response.data.access_token;
    })
    .catch((error) => {
      console.error("Error retrieving access token: " + error);
      throw error; // Re-throw the error so it can be caught in your calling code
    });
};


/**
 * CORS and Express Setup
 */
app.use(cors());
app.use(express.json());
const port = process.env.PORT || 80;

/**
 * Check if server is up
 */
app.get("/", (req, res) => {
  res.send("Hello, Happy Server!");
});

/**
 * Get List of All users from GRAPH API using Access Token
 */
app.get("/users", async (req, res) => {
  const applicationToken = await getAccessToken();
  axios
    .get("https://graph.microsoft.com/v1.0/users", {
      headers: {
       // Authorization: `Bearer ${req.body.accessToken} `,
        Authorization: `Bearer ${applicationToken} `,
      },
    })
    .then((response) => {
      res.json({ users: response.data });
    })
    .catch((error) => {
      // Handle errors here
      if (error.response) {
        console.error("Response status:", error.response.status);
        console.error("Response data:", error.response.data);
        res.status(error.response.status).json({ error: error.response.data });
      } else {
        console.error("Error:", error.message);
        res.status(500).json({ error: "Internal Server Error" });
      }
    });
});


/**
 * Send Notification for Happy Companies
 *
 */

app.post("/v1/api/send-notification", async (req, res) => {
  try {
    const applicationToken = await getAccessToken();
    console.log("applicationToken = ",applicationToken);
    const Decoded_token = jwtDecode(applicationToken);
    console.log("Decoded_token = ",Decoded_token)
    const data = JSON.parse(process.env.AdminAPPID);
    console.log("Data = ", data);
    console.log("Tid = ", Decoded_token.tid);
    //const AppID = data?.[Decoded_token.tid];
    const AppID = "60a0cac6-1c81-4df2-b558-8b6c4d15f711";
    console.log("App ID = ", AppID);
    
    const notificationPromises = req.body.userIds.map((userId) => {
      const postData = {
        topic: {
          source: "text",
          value: "Hello World",
          webUrl: `https://teams.microsoft.com/l/entity/${AppID}/a7fc217d-e27c-4fa3-b069-2e1055bb7710?tenantId=${Decoded_token.tid}`,
        },
        activityType: "appluaseCreated",
        previewText: {
          content: "Happy Companies > Home",
        }
      };

      return axios
        .post(
          `https://graph.microsoft.com/v1.0/users/${userId}/teamwork/sendActivityNotification`,
          postData,
          {
            headers: {
              "Content-Type": "application/json",
              Authorization: "bearer " + applicationToken,
            },
          }
        )
        .then((response) => {
          console.log(
            `Notification sent to user ${userId}, statusCode: ${response.status}`
          );
        })
        .catch((error) => {
          console.log(`Error sending notification to user ${userId}: ${error}`);
        });
    });

    // Wait for all notification promises to resolve
    Promise.all(notificationPromises)
      .then(() => {
        res.sendStatus(200);
      })
      .catch((error) => {
        console.log("Bulk notification sending error: " + error);
        res.sendStatus(400);
      });
  } catch (error) {
    // Handle verification or decoding errors here
    console.error(error);
    res.sendStatus(400);
  }
});

/**
 * Start Server
 */
app.listen(port, () => {
  console.log("Server is running");
});
