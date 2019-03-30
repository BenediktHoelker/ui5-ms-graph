/* global QUnit */
QUnit.config.autostart = false;

sap.ui.getCore().attachInit(function () {
	"use strict";

	sap.ui.require([
		"com/iot/ui5-ms-graph/test/unit/AllTests"
	], function () {
		QUnit.start();
	});
});