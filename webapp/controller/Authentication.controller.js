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
        var sessionModel = this.getOwnerComponent().getModel("session");
        var eventsModel = this.getOwnerComponent().getModel();
        sessionModel.setProperty(
          "/currentDateTime",
          new Date("2017", "03", "10", "0", "0")
        );
        sessionModel.setProperty("/", [
          {
            start: new Date("2017", "03", "10", "0", "0"),
            end: new Date("2017", "05", "16", "23", "59"),
            title: "Vacation",
            info: "out of office",
            type: "Type04",
            tentative: false
          }
        ]);
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
      },

      onQuery: function() {
        var eventsModel = this.getOwnerComponent().getModel();
        eventsModel.setProperty("/", [
          {
            start: new Date("2017", "0", "8", "08", "30"),
            end: new Date("2017", "0", "8", "09", "30"),
            title: "Meet Max Mustermann",
            type: "Type02",
            tentative: false
          }
        ]);
      },

      queryEvents: function() {
        var eventsModel = this.getOwnerComponent().getModel();
        var sessionModel = this.getOwnerComponent().getModel("session");
        var accessToken = sessionModel.getProperty("/token");
        fetch(applicationConfig.graphEndpoint + "/calendar/events", {
          method: "GET", // or 'PUT'
          headers: { Authorization: "Bearer " + accessToken }
        })
          .then(res => res.json())
          .then(response =>
            response.value
              .filter(event => event.start.dateTime && event.end.dateTime)
              .map(event => {
                return {
                  ...event,
                  start: new Date(event.start.dateTime),
                  end: new Date(event.end.dateTime)
                };
              })
          )
          .then(events =>
            eventsModel.setProperty("/", events)
          );
      }
    });
  }
);
