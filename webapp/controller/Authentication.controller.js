sap.ui.define(
  [
    "com/iot/ui5-ms-graph/msal/config",
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel"
  ],
  function(applicationConfig, Controller, JSONModel) {
    "use strict";

    var myMSALObj = new Msal.UserAgentApplication(
      applicationConfig.clientID,
      applicationConfig.authority,
      acquireTokenRedirectCallBack,
      { storeAuthStateInCookie: true, cacheLocation: "localStorage" }
    );

    function acquireTokenRedirectCallBack(errorDesc, token, error, tokenType) {
      if (tokenType === "access_token") {
        callMSGraph(applicationConfig.graphEndpoint, token, graphAPICallback);
      } else {
        console.log("token type is:" + tokenType);
      }
    }

    function signOut() {
      myMSALObj.logout();
    }

    return Controller.extend("com.iot.ui5-ms-graph.controller.InitialView", {
      onInit: function() {
        var sessionModel = new JSONModel();
        this.getOwnerComponent().setModel(sessionModel, "session");
      },

      signIn: function() {
        var sessionModel = this.getOwnerComponent().getModel("session");

        myMSALObj.loginPopup(applicationConfig.graphScopes).then(
          function(idToken) {
            myMSALObj
              .acquireTokenSilent(applicationConfig.graphScopes)
              .then(function(accessToken) {
                sessionModel.setProperty("/token", accessToken);
                return accessToken;
              })
              .then(function(accessToken) {
                return fetch(applicationConfig.graphEndpoint, {
                  method: "GET", // or 'PUT'
                  headers: { Authorization: "Bearer " + accessToken }
                });
              })
              .then(res => res.json())
              .then(response =>
                sessionModel.setProperty("/userData", response)
              );
          },
          function(error) {
            console.log(error);
          }
        );
      },

      signOut: function() {
        signOut();
      }
    });
  }
);
