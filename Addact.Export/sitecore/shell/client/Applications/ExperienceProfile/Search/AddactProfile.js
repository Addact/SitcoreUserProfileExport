
var Addact = {
    ExportData: function () {
        var apiurl = "/api/bghExportData/ExportProfile";
        var query = "";
        var fromdate = $("[data-sc-id=FromDatePick]").val();
        var todate = $("[data-sc-id=ToDatePick]").val();
        if (fromdate != undefined && fromdate != "") {
            query += "?startDate=" + fromdate;

        }
        if (todate != undefined && fromdate != "") {
            if (query != "") {
                query += "&endDate=" + todate;
            }
            else { query += "?endDate=" + todate; }
        }
        window.location.assign(apiurl + query);
    },
    initialized: function () {
        var buttonhtml = '<button class="btn sc-button btn-default noText sc_Button_68 " title="Export Profile Data" onclick="javascript:Addact.ExportData()"  type="button">';
        buttonhtml += '<div class="sc-icon data-sc-registered" style="background-position: center center; background-image: url(&quot;/sitecore/shell/themes/standard/~/icon/Office/16x16/box_out.png&quot;);"></div>';
        buttonhtml += '<span class="sc-button-text data-sc-registered"></span></button>';
        if ($(".sc-applicationHeader-contextSwitcher").length > 0) {


            $(".sc-applicationHeader-contextSwitcher").html(buttonhtml);
            $(".sc-applicationHeader-actions").css("width", "80%");

        }


    },

}


$(function () { Addact.initialized(); });