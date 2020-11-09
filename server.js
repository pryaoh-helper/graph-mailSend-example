

// Load environment variables from project .env file
require('node-env-file')(__dirname + '/.env');

const debug = require("debug")("mail");

const express = require('express');
const app = express();

const request = require('request');
const https = require('https');


const tenantId = process.env.TENANT_ID || "e401ddf3-e2e4-4248-bcac-a5c248130505";
const clientId = process.env.CLIENT_ID || "6d758728-607f-437c-99b4-4702a91a3bcd";
const clientSecret = process.env.CLIENT_SECRET || "9.688H3_8h38Jwj5-p-.Cr1P_shWdo767l";
const scopes = process.env.SCOPES || "https://graph.microsoft.com/.default"; 

const port = process.env.PORT || 8080;
let redirectURI = process.env.REDIRECT_URI
if (!redirectURI) {
   // domain hosting
   if (process.env.PROJECT_DOMAIN) {
      redirectURI = "https://" + process.env.PROJECT_DOMAIN + "/permissions";
   }
   else {
      // defaults to localhost
      redirectURI = `http://localhost:${port}/permissions`;
   }
}


const state = process.env.STATE || "GSENC_TEST";
const consentUrl = `https://login.microsoftonline.com/${tenantId}/adminconsent?`
    + `&client_id=${clientId}`
    + `&state=${state}`
    + `&redirect_uri=${redirectURI}`;

const read = require("fs").readFileSync;
const join = require("path").join;
const str = read(join(__dirname, '/www/index.ejs'), 'utf8');
const ejs = require("ejs");
const compiled = ejs.compile(str)({ "link": consentUrl }); // inject the link into the template

/// 라우팅
app.get("/index.html", function (req, res) {
    debug("serving the html page (generated from an EJS template)");
    res.send(compiled);
 });


 app.get("/", function (req, res) {
    res.redirect("/index.html");
 });

     
 /// 호스팅
const path = require('path');
const { data } = require('node-env-file');
app.use("/", express.static(path.join(__dirname, 'www')));


app.get("/permissions", function (req, res) {
    debug("permissions callback entered");

    if (req.query.error) {
        // 에러 처리
        debug("received err: " + req.query.error);
        res.send(`<h1>인증이 실패하였습니다.</h1><p>${req.query.error_description}</p>`);
        return;
    }


     // Check request parameters correspond to the spec
   if (!req.query.state) {
    debug("expected state query parameters are not present");
    res.send("<h1>인증이 실패하였습니다.</h1><p>State not present</p>");
    return;
   }

    // Check State 
    // [NOTE] we implement a Security check below, but the State variable can also be leveraged for Correlation purposes
    if (state != req.query.state) {
        debug("State does not match");
        res.send("<h1>인증이 실패하였습니다.</h1><p>State does not match</p>");
        return;
    }
    
    if(!req.query.admin_consent)
    {
        debug("consent set to false");
        res.send("<h1>인증이 실패하였습니다.</h1><p>Consent must be true.</p>");
        return;
    }

    onPermissionCallback(res);
});

function onPermissionCallback(res) {
     // on Permission callback

     const formData = {
        grant_type: 'client_credentials',
        client_id: clientId,
        client_secret: clientSecret,
        scope: scopes,        
    };


    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const options = {
        method: "POST",
        url: url,
        headers: {
           "content-type": "application/x-www-form-urlencoded"
        },
        form: formData
     };

     request(options, function (error, response, body) {
        if (error) {            
            res.send("<h1>AzureAD OAuth could not complete</h1><p>Sorry, could not retreive your access token. Try again...</p>");
            return;
        }

        if (response.statusCode != 200) {
            debug("access token not issued with status code: " + response.statusCode);
            res.send("<h1>AzureAD OAuth could not complete</h1><p>Sorry, could not retreive your access token. Try again...</p>");
        }


        // Check payload
        const json = JSON.parse(body);

        debug("AzureAD OAuth completed, fetched tokens: " + JSON.stringify(json))

        const str = read(join(__dirname, '/www/sendmail.ejs'), 'utf8');
        const compiled = ejs.compile(str)({ "accessToken": json.access_token });
        res.send(compiled);

     });
     
      
 }

 // listen
app.listen(port, function () {
    console.log("Graph-MailSend-Example started on port: " + port);
 });