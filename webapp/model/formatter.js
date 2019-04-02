sap.ui.define([], function () {
	"use strict";

	return {
		calendarDayType: function (category) {
            switch (category){
                case "Frankonia": 
                    return "NonWorking";
                case "Manufactum": 
                    return "Type02";
                case "Intorq": 
                    return "Type03";
                case "Esprit": 
                    return "Type04";
            }
		}
    }
});