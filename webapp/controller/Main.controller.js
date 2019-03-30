/* global msalconfig, Msal */
sap.ui.define(["sap/ui/core/mvc/Controller", "sap/m/MessageToast", "sap/ui/model/json/JSONModel"],
  function (Controller, MessageToast, JSONModel){
    onInit: function () {
        this.oUserAgentApplication = new Msal.UserAgentApplication(msalconfig.clientID, null,
          function (errorDesc, token, error, tokenType) {
            if (errorDesc) {
              var formattedError = JSON.stringify(error, null, 4);
              if (formattedError.length < 3) {
                formattedError = error;
              }
              MessageToast.show("Error, please check the $.sap.log for details");
              $.sap.log.error(error);
              $.sap.log.error(errorDesc);
            } else {
              this.fetchUserInfo();
            }
          }.bind(this), {
            redirectUri: msalconfig.redirectUri
          });
        //Previous version of msal uses redirect url via a property
        if (this.oUserAgentApplication.redirectUri) {
          this.oUserAgentApplication.redirectUri = msalconfig.redirectUri;
        }
        // If page is refreshed, continue to display user info
        if (!this.oUserAgentApplication.isCallback(window.location.hash) && window.parent === window) {
          var user = this.oUserAgentApplication.getUser();
          if (user) {
            this.fetchUserInfo();
          }
        }
      },

      //************* MSAL functions *****************//
onSwitchSession: function (oEvent) {
    var oSessionModel = oEvent.getSource().getModel('session');
    var bIsLoggedIn = oSessionModel.getProperty('/displayName');
    if (bIsLoggedIn) {
        this.oUserAgentApplication.logout();
        return;
    }
    this.fetchUserInfo();
    },


    fetchUserInfo: function () {
        this.callGraphApi(msalconfig.graphBaseEndpoint + msalconfig.userInfoSuffix, function (response) {
            $.sap.log.info("Logged in successfully!", response);
            this.getView().getModel('session').setData(response);
        }.bind(this));
        }
    })