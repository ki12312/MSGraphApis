require("dotenv").config();
const axios = require("axios");
const { jwtDecode } =require("jwt-decode");

async function authenticateToken(req, res, next) {
  try {
    const clientid = process.env.CLIENT_ID;
    const clientsecret = process.env.CLIENT_SECRET;
    const scopes = "https://graph.microsoft.com/.default";
    const token = req.header("Authorization");



    if (!token) {
      return res.status(400).json({ error: "No token provided" });
    }

    const url = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

    const params = new URLSearchParams();
    params.append("client_id", clientid);
    params.append("client_secret", clientsecret);
    params.append("grant_type", "urn:ietf:params:oauth:grant-type:jwt-bearer");
    params.append("assertion", token);
    params.append("requested_token_use", "on_behalf_of");
    params.append("scope", scopes);

    const response = await axios.post(url, params, {
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });

    if (response.status === 200) {
      const accessToken = response.data.access_token;

      const payload = jwtDecode(accessToken);
      req.body.accessToken = accessToken;
      req.body.oid = payload.oid;
      req.body.tid = payload.tid;

      axios
        .get(
          "https://graph.microsoft.com/beta/me/settings/regionalAndLanguageSettings",
          {
            headers: {
              Authorization: `Bearer ${req.body.accessToken} `,
            },
          }
        )
        .then((response) => {
          const language = response?.data?.defaultDisplayLanguage?.locale || "en-US";
          if (language === "en-US") req.body.language = "en";
          else req.body.language = "fr";

          next();
        })
        .catch((error) => {
          // Handle errors here
          if (error.response) {
            console.error("Response status:", error.response.status);
            console.error("Response data:", error.response.data);
            return res.status(400).json({ error });
          } else {
            console.error("Error:", error.message);
            return res.status(400).json({ error: error.message });
          }
        });
    } else {
      return res.status(400).json({ error: response.data.error });
    }
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "An error occurred" });
  }
}

module.exports = authenticateToken;
