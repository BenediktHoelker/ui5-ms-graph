sap.ui.define(["sap/ui/core/mvc/Controller"], function(Controller) {
  "use strict";

  var applicationConfig = {
    clientID: "94bb3d16-4be3-4020-b929-ec0ca5a4103a",
    authority: "https://login.microsoftonline.com/organizations",
    graphScopes: ["user.read"],
    graphEndpoint: "https://graph.microsoft.com/v1.0/me"
  };

  var myMSALObj = new Msal.UserAgentApplication(
    applicationConfig.clientID,
    applicationConfig.authority,
    acquireTokenRedirectCallBack,
    { storeAuthStateInCookie: true, cacheLocation: "localStorage" }
  );

  function signIn() {
    myMSALObj.loginPopup(applicationConfig.graphScopes).then(
      function(idToken) {
        //Login Success
        acquireTokenPopupAndCallMSGraph();
      },
      function(error) {
        console.log(error);
      }
    );
  }

  function signOut() {
    myMSALObj.logout();
  }

  function acquireTokenPopupAndCallMSGraph() {
    //Call acquireTokenSilent (iframe) to obtain a token for Microsoft Graph
    myMSALObj.acquireTokenSilent(applicationConfig.graphScopes).then(
      function(accessToken) {
        // new sap.ui.model.odata.v2.ODataModel("https://graph.microsoft.com/v1.0", {
        //  headers: { Authorization: "Bearer " + accessToken }
        // });
        sap.ui.getCore().BenesModel = new sap.ui.model.odata.v4.ODataModel({
          earlyRequests: true,
          serviceUrl: "https://graph.microsoft.com/v1.0/me/",
          synchronizationMode: "None"
        });
        callMSGraph(
          applicationConfig.graphEndpoint,
          accessToken,
          graphAPICallback
        );
      },
      function(error) {
        console.log(error);
        // Call acquireTokenPopup (popup window) in case of acquireTokenSilent failure due to consent or interaction required ONLY
        if (
          error.indexOf("consent_required") !== -1 ||
          error.indexOf("interaction_required") !== -1 ||
          error.indexOf("login_required") !== -1
        ) {
          myMSALObj.acquireTokenPopup(applicationConfig.graphScopes).then(
            function(accessToken) {
              callMSGraph(
                applicationConfig.graphEndpoint,
                accessToken,
                graphAPICallback
              );
            },
            function(error) {
              console.log(error);
            }
          );
        }
      }
    );
  }

  function callMSGraph(theUrl, accessToken, callback) {
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.onreadystatechange = function() {
      if (this.readyState == 4 && this.status == 200)
        callback(JSON.parse(this.responseText));
    };
    xmlHttp.open("GET", theUrl, true); // true for asynchronous
    xmlHttp.setRequestHeader("Authorization", "Bearer " + accessToken);
    xmlHttp.send();
  }

  function graphAPICallback(data) {
    //Display user data on DOM
    console.log(JSON.stringify(data));
  }

  function acquireTokenRedirectCallBack(errorDesc, token, error, tokenType) {
    if (tokenType === "access_token") {
      callMSGraph(applicationConfig.graphEndpoint, token, graphAPICallback);
    } else {
      console.log("token type is:" + tokenType);
    }
  }

  return Controller.extend("com.iot.ui5-ms-graph.controller.InitialView", {
    onInit: function() {},
    signIn: function() {
      signIn();
    },

    signOut: function() {
      signOut();
    }
  });
});
