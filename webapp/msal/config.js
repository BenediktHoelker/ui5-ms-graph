sap.ui.define([], function() {
  "use strict";

  return {
    clientID: "94bb3d16-4be3-4020-b929-ec0ca5a4103a",
    authority: "https://login.microsoftonline.com/organizations",
    graphScopes: ["user.read", "calendars.read", "calendars.ReadWrite"],
    graphEndpoint: "https://graph.microsoft.com/v1.0/me"
  };
});
