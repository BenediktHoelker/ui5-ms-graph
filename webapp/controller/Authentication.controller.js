sap.ui.define(
  [
    "com/iot/ui5-ms-graph/msal/config",
    "sap/ui/core/mvc/Controller",
    "com/iot/ui5-ms-graph/model/formatter"
  ],
  function(applicationConfig, Controller, formatter) {
    "use strict";

    var myMSALObj = new Msal.UserAgentApplication(
      applicationConfig.clientID,
      applicationConfig.authority,
      null,
      { storeAuthStateInCookie: true, cacheLocation: "localStorage" }
    );

    return Controller.extend("com.iot.ui5-ms-graph.controller.InitialView", {
      formatter: formatter,
      onInit: function() {
        var date = new Date();

        date.setDate(date.getDate() - 14);

        var sessionModel = this.getOwnerComponent().getModel("session");
        var eventsModel = this.getOwnerComponent().getModel();
        sessionModel.setProperty("/startDate", date);
        this._signIn().then(() => {
          var oDateRange = this._getCalendarDateRange();
          this._queryEvents(oDateRange);
        });
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

      onStartDateChange: function(oEvent) {
        var oDateRange = this._getCalendarDateRange();
        this._queryEvents(oDateRange);
      },

      onEditPress: function(oEvent) {
        var oObject = oEvent
          .getSource()
          .getBindingContext()
          .getObject();
        var sId = oObject.id;
        var accessToken = this.getView()
          .getModel("session")
          .getProperty("/token");

        this._updateEvent(
          sId,
          { subject: "[APP] " + object.subject },
          accessToken
        );
      },

      _updateEvent: function(eventId, update, accessToken) {
        return fetch(applicationConfig.graphEndpoint + `/events/${eventId}`, {
          method: "PATCH", // or 'PUT'
          headers: {
            "Content-Type": "application/json",
            Authorization: "Bearer " + accessToken
          },
          body: JSON.stringify(update)
        });
      },

      onCategorySelect: function(oEvent) {
        var oListItem = oEvent.getSource();
        // var sSelectedKey = oListItem.getSelectedKey();
        var oObject = oListItem.getBindingContext().getObject();

        var sId = oObject.id;
        var accessToken = this.getView()
          .getModel("session")
          .getProperty("/token");

        this._updateEvent(
          sId,
          { categories: oObject.categories },
          accessToken
        ).then(() => {
          // oObject.categories = [sSelectedKey];
          // this.getView()
          //   .getModel()
          //   .refresh(true);
        });
      },

      navToTable: function() {
        var oNavContainer = this.byId("navContainer");
        var oTablePage = this.byId("tablePage");
        oNavContainer.to(oTablePage);
      },

      navToCalendar: function() {
        var oNavContainer = this.byId("navContainer");
        var oCalendarPage = this.byId("calendarPage");
        oNavContainer.back(oCalendarPage);
      },

      _getCalendarDateRange: function() {
        var oCalendar = this.byId("PC1");
        var oDateRange = oCalendar._getFirstAndLastRangeDate();
        var oStartDate = oDateRange.oStartDate._oUDate.oDate.toISOString();
        var oEndDate = oDateRange.oEndDate._oUDate.oDate.toISOString();
        return { startDate: oStartDate, endDate: oEndDate };
      },

      _signIn: function() {
        var sessionModel = this.getOwnerComponent().getModel("session");

        return myMSALObj
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

      _queryEvents: function(dateRange) {
        var eventsModel = this.getOwnerComponent().getModel();
        var sessionModel = this.getOwnerComponent().getModel("session");
        var accessToken = sessionModel.getProperty("/token");
        var sQuery = `/calendarview?startdatetime=${
          dateRange.startDate
        }&enddatetime=${dateRange.endDate}&$top=200`;
        return fetch(applicationConfig.graphEndpoint + sQuery, {
          method: "GET", // or 'PUT'
          headers: { Authorization: "Bearer " + accessToken }
        })
          .then(res => res.json())
          .then(response =>
            response.value
              // .filter(event => event.start.dateTime && event.end.dateTime)
              .map(event => {
                return {
                  ...event,
                  start: new Date(event.start.dateTime),
                  end: new Date(event.end.dateTime)
                };
              })
          )
          .then(events => eventsModel.setProperty("/", events));
      }
    });
  }
);
