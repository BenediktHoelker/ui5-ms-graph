sap.ui.define(
  [
    "com/iot/ui5-ms-graph/msal/config",
    "sap/ui/core/mvc/Controller",
  ],
  function(applicationConfig, Controller) {
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
				this._signIn();
      },

      signIn: function() {
        myMSALObj.loginPopup(applicationConfig.graphScopes).then(
          function() {
						this._signIn();
					},
          function(error) {
            console.log(error);
          }
        );
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
      },

      signOut: function() {
        signOut();
      }
    });
  }
);
