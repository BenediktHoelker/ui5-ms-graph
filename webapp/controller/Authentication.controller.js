sap.ui.define(
  ["com/iot/ui5-ms-graph/msal/config", "sap/ui/core/mvc/Controller"],
  function(applicationConfig, Controller) {
    "use strict";

    var myMSALObj = new Msal.UserAgentApplication(
      applicationConfig.clientID,
      applicationConfig.authority,
      null,
      { storeAuthStateInCookie: true, cacheLocation: "localStorage" }
    );

    return Controller.extend("com.iot.ui5-ms-graph.controller.InitialView", {
      onInit: function() {
        this._signIn();
      },

      onSwitchSession: function() {
        var sessionModel = this.getOwnerComponent().getModel("session");
        var isLoggedIn = sessionModel.getProperty("/userData/displayName");
        if (isLoggedIn) {
          myMSALObj.logout();
        } else {
          myMSALObj
            .loginPopup(applicationConfig.graphScopes)
            .then(() => this._signIn())
            .catch(error => console.log(error));
        }
      },

      _signIn: function() {
        var sessionModel = this.getOwnerComponent().getModel("session");

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
          .then(response => sessionModel.setProperty("/userData", response));
      }
    });
  }
);
