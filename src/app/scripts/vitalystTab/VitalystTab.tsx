import * as React from "react";
import ScrollableAnchor from "react-scrollable-anchor";

import {
    PrimaryButton,
    TeamsThemeContext,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface,
    getContext
} from "msteams-ui-components-react";

import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import $ = require("jquery");
import "bootstrap";
import Popper from "popper.js";
import { url } from "inspector";
import { AppBasedLinkQuery } from "botframework-connector/lib/teams/models/mappers";
import { MsTeamsApiRouter } from "express-msteams-host";
import { SkypeMentionNormalizeMiddleware, AutoSaveStateMiddleware, ConversationState } from "botbuilder";

import {
    Modal,
    Header,
    Footer,
    Button,
    Body,
    Title
} from "react-bootstrap-modal";
import e = require("express");
import { CSSProperties } from "typestyle/lib/types";
import * as CSS from "csstype";

// tslint:disable-next-line:interface-name
interface Style extends CSS.Properties, CSS.PropertiesHyphen {}

/**
 * State for the vitalystTabTab React component
 */
export interface IVitalystTabState extends ITeamsBaseComponentState {
    entityId?: string;
}

/**
 * Properties for the vitalystTabTab React component
 */
export interface IVitalystTabProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the Vitalyst content page
 */
export class VitalystTab extends TeamsBaseComponent<IVitalystTabProps, IVitalystTabState> {
    public handleClick = (btnID: number, strModalURL: string, modalType: string, videoTitle: string, videoApp: string, videoDesc: string) => {
        if (modalType === "enrollment") {
            const webinarID = strModalURL.slice(-5);
            $("#submitEventId").val(webinarID);
            $("#submitEventId").attr("readonly", "true");
            $("#submitTitle").html(videoTitle);
            const eventId = $("#submitEventId").val();
            // tslint:disable-next-line:only-arrow-functions
            this.ebAPIRequest("GET", "https://vitalyst.eventbuilder.com/api/event/" + eventId, function(err: any, result: any) {
                if (err) {
                    return console.error(err);
                }
                const occurrenceId = result.occurrence.id;
                $("#submitOccurrenceId").val(occurrenceId);
                $("#submitOccurrenceId").attr("readonly", "true");
            }, "");
            $("#modalSubmit").modal();
        } else if (modalType === "video") {
            $("#modalVideo").prop("src", strModalURL);
            $("#videoTitle").html(videoTitle);
            $("#videoApp").html(videoApp);
            $("#videoDesc").html(videoDesc);
            $("#modalVideoDialog").modal();
        }
    }
    public submitReg = () => {
        const eventId = $("#submitEventId").val();
        const occurrenceId = $("#submitOccurrenceId").val();
        const outputJSON = "{\"firstname\":\"" + $("#submitFName").val() + "\",\"lastname\":\"" + $("#submitLName").val() + "\",\"email\":\"" + $("#submitAddress").val() + "\",\"questions\":[]}";
        // tslint:disable-next-line:only-arrow-functions
        this.ebAPIRequest("POST", "https://vitalyst.eventbuilder.com/api/event/" + eventId + "/occurrence/" + occurrenceId + "/registrants", outputJSON, function(err: any, result: any) {
            if (err) { return alert(err); }
            $("#submitEventId").val("");
            $("#submitOccurrenceId").val("");
            $("#submitFName").val("");
            $("#submitLName").val("");
            $("#submitAddress").val("");
            $("#modalSubmit").modal("hide");
        });
    }
    public ebAPIRequest = (method: string, strUrl: string, body: any, cb: any) => {
        if (typeof body === "function") { cb = body; body = null; }
        const xhr = new XMLHttpRequest();
        xhr.open(method, strUrl, true);
        xhr.setRequestHeader("Content-Type", "application/json");
        // tslint:disable-next-line:only-arrow-functions
        xhr.onreadystatechange = function() {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                let response = xhr.response;
                // tslint:disable-next-line:only-arrow-functions
                const headers = xhr.getAllResponseHeaders().split("\r\n").reduce(function(acc, current, i) {
                    const parts = current.split(": ");
                    acc[parts[0]] = parts[1];
                    return acc;
                }, {});

                try {
                    response = JSON.parse(xhr.response);
                } catch (err) {
                    response = xhr.response;
                }
                if (xhr.status === 200) {
                    cb(null, response, headers);
                } else {
                    cb(xhr.statusText, response, headers);
                }
            }
        };
        xhr.send(body);
    }
    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                this.setState({
                    entityId: context.entityId
                });
            });
        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
            });
        }
    }
    public componentDidMount() {
        // tslint:disable-next-line:only-arrow-functions
        $("#filClearAll").on("click", function() {
            $("input[id*='chk']").prop("checked", false);
            $("[class^='col-sm-4 d-flex pb-3']").attr("style", "display: '' !important");
        });
        // tslint:disable-next-line:only-arrow-functions
        $("#filQClearAll").on("click", function() {
            $("input[id*='chk']").prop("checked", false);
            $("[class^='col-sm-4 d-flex pb-3']").attr("style", "display: '' !important");
        });
        // tslint:disable-next-line:only-arrow-functions
        $("#filClearAllNav").on("click", function() {
            $("input[id*='chk']").prop("checked", false);
            $("[class^='col-sm-4 d-flex pb-3']").attr("style", "display: '' !important");
        });
        // tslint:disable-next-line:only-arrow-functions
        $("#filQA365").on("click", function() {
            $("[class='col-sm-4 d-flex pb-3']").attr("style", "display: none !important");
            $("span[class='a365']:contains(Active365:true)").parent().parent().parent().parent().attr("style", "display: '' !important");
        });
        // tslint:disable-next-line:only-arrow-functions
        $("#myBtn").on("click", function() {
            $("#modalSubmit").modal();
        });
        // tslint:disable-next-line:only-arrow-functions
        $("#closeVideoDialog").on("click", function() {
            $("#modalVideoDialog").modal("hide");
            const vidSrc: any = $("#modalVideo").attr("src");
            $("#modalVideo").attr("src", "");
        });
        // tslint:disable-next-line:only-arrow-functions
        $("#btnND").on("click", function() {
            $("#modalAppDialog").modal();
        });
        // tslint:disable-next-line:only-arrow-functions
        $("#closeAppDialog").on("click", function() {
            $("#modalAppDialog").modal("hide");
        });
        // tslint:disable-next-line:only-arrow-functions
        $.expr[":"].contains = $.expr.createPseudo(function(arg) {
            // tslint:disable-next-line:only-arrow-functions
            return function( elem ) {
                return $(elem).text().toUpperCase().indexOf(arg.toUpperCase()) >= 0;
            };
        });
        if ($("#webEnabled").val() === "false") {
            $("#webinarsLink").addClass("hidden");
        }
        if ($("#vidEnabled").val() === "false") {
            $("#videosLink").addClass("hidden");
        }
        if ($("#custEnabled").val() === "false") {
            $("#customLink").addClass("hidden");
        }
        if ($("#defaultTabName").val() === "web") {
            $("#webinarsHeader").removeClass("hidden");
            $("#webinars").removeClass("hidden");
            $("#webinarsNav").addClass("active");
        }
        if ($("#defaultTabName").val() === "vid") {
            $("#videosHeader").removeClass("hidden");
            $("#videos").removeClass("hidden");
            $("#videosNav").addClass("active");
        }
        if ($("#defaultTabName").val() === "cust") {
            $("#customHeader").removeClass("hidden");
            $("#custom").removeClass("hidden");
            $("#customNav").addClass("active");
        }
        if ($("#showFilterList").val() === "false") {
            $("#filterDropdown").addClass("hidden");
        }
        if ($("#showClearButton").val() === "false") {
            $("#clearFilters").addClass("hidden");
        }
        // tslint:disable-next-line:only-arrow-functions
        $("#searchbox").on("keyup", function() {
            $("[class='col-sm-4 d-flex pb-3']").attr("style", "display: none !important");
            const filter = $(this).val();
            $("h5[class='card-title']:contains(" + filter + ")").parent().parent().parent().attr("style", "display: '' !important");
            $("h6[class='card-subtitle']:contains(" + filter + ")").parent().parent().parent().attr("style", "display: '' !important");
            $("span[class='app-text']:contains(" + filter + ")").parent().parent().parent().parent().attr("style", "display: '' !important");
          });
        // tslint:disable-next-line:only-arrow-functions
        $("#searchbox").on("input", function() {
            const filter = $(this).val();
            if (filter === "") {
                $("[class^='col-sm-4 d-flex pb-3']").attr("style", "display: '' !important");
            }
        });
        $(".next").on("click", function() {
        const currentPanel: any = $(this).closest("div");
        const nextPanel: any = currentPanel.find("div");
        if (nextPanel.attr("class").indexOf("hidden") !== 1) {
            nextPanel.removeClass("hidden");
            $(this).css("display", "none");
            $(this).parent().css("display", "none");
        }
        return false;
        });
        // tslint:disable-next-line:only-arrow-functions
        $("#testButton").on("click", function() {
            $("#formRow").removeClass("hidden");
        });
        // tslint:disable-next-line:only-arrow-functions
        $("#cancelForm").on("click", function() {
            $("#submitEventId").val("");
            $("#submitOccurrenceId").val("");
            $("#submitFName").val("");
            $("#submitLName").val("");
            $("#submitAddress").val("");
            $("#modalSubmit").modal("hide");
        });
        // tslint:disable-next-line:only-arrow-functions
        $("#webinarsLink").on("click", function() {
            $("#webinars").removeClass("hidden");
            $("#webinarsHeader").removeClass("hidden");
            $("#webinarsNav").addClass("active");
            $("#videos").addClass("hidden");
            $("#videosHeader").addClass("hidden");
            $("#videosNav").removeClass("active");
            $("#custom").addClass("hidden");
            $("#customHeader").addClass("hidden");
            $("#customNav").removeClass("active");
            $("#filterQtrDropdown").addClass("hidden");
        });
        // tslint:disable-next-line:only-arrow-functions
        $("#videosLink").on("click", function() {
            $("#webinars").addClass("hidden");
            $("#webinarsHeader").addClass("hidden");
            $("#webinarsNav").removeClass("active");
            $("#videos").removeClass("hidden");
            $("#videosHeader").removeClass("hidden");
            $("#videosNav").addClass("active");
            $("#custom").addClass("hidden");
            $("#customHeader").addClass("hidden");
            $("#customNav").removeClass("active");
            if ($("#showReleaseList").val() === "true") {
                $("#filterQtrDropdown").removeClass("hidden");
            } else {
                $("#filterQtrDropdown").addClass("hidden");
            }
        });
        // tslint:disable-next-line:only-arrow-functions
        $("#customLink").on("click", function() {
            $("#webinars").addClass("hidden");
            $("#webinarsHeader").addClass("hidden");
            $("#webinarsNav").removeClass("active");
            $("#videos").addClass("hidden");
            $("#videosHeader").addClass("hidden");
            $("#videosNav").removeClass("active");
            $("#custom").removeClass("hidden");
            $("#customHeader").removeClass("hidden");
            $("#customNav").addClass("active");
            $("#filterQtrDropdown").addClass("hidden");
        });
        let strFilterList: any = "";
        strFilterList = $("#appFilters").val();
        // tslint:disable-next-line:only-arrow-functions
        $.each(strFilterList.split(";").reverse(), function(index, item) {
            $("#dropdown-menu").prepend("<a class='dropdown-item' href='#' id='fil" + item.replace(" ", "") + "'>" + item + "</a>");
            // tslint:disable-next-line:only-arrow-functions
            $("#fil" + item.replace(" ", "")).bind("click", function() {
                $("[class='col-sm-4 d-flex pb-3']").attr("style", "display: none !important");
                $("h5[class='card-title']:contains('" + item + "')").parent().parent().parent().attr("style", "display: '' !important");
                $("h6[class='card-subtitle']:contains('" + item + "')").parent().parent().parent().attr("style", "display: '' !important");
                $("span[class='app-text']:contains('" + item + "')").parent().parent().parent().parent().attr("style", "display: '' !important");
            });
        });
        let strQtrList: any = "";
        strQtrList = $("#qtrFilters").val();
        if (strQtrList !== "") {
            // tslint:disable-next-line:only-arrow-functions
            $.each(strQtrList.split(";").reverse(), function(qtrIndex, item) {
                $("#filterForm").prepend("<div class='form-check'>&nbsp;&nbsp;&nbsp;<input type='checkbox' id='chk" + item.replace(" ", "") + "'></input>&nbsp;&nbsp;<label>" + item.replace(" ", "") + "</label></div>");
                // tslint:disable-next-line:only-arrow-functions
                $("#chk" + item.replace(" ", "")).bind("click", function() {
                    $("[class='col-sm-4 d-flex pb-3']").attr("style", "display: none !important");
                    const varQtrs = strQtrList.split(";").reverse();
                    let bolAtLeastOne: boolean = false;
                    // tslint:disable-next-line:only-arrow-functions
                    $.each(varQtrs, function(index, qtrItem) {
                        if ($("#chk" + qtrItem.replace(" ", "")).prop("checked") === true) {
                            $("p[class='card-text']:contains('" + qtrItem + "')").parent().parent().parent().attr("style", "display: '' !important");
                            bolAtLeastOne = true;
                        }
                    });
                    if (bolAtLeastOne === false) {
                        $("[class^='col-sm-4 d-flex pb-3']").attr("style", "display: '' !important");
                    }
                });
            });
        }
        this.componentDidUpdate();
    }
    public componentDidUpdate() {
        if ($("#defaultTabName").val() === "web") {
            $("#filterQtrDropdown").addClass("hidden");
            $("#webinarsLink").click();
        }
        if ($("#defaultTabName").val() === "vid") {
            if ($("#showReleaseList").val() === "true") {
                $("#filterQtrDropdown").removeClass("hidden");
            } else {
                $("#filterQtrDropdown").addClass("hidden");
            }
            $("#videosLink").click();
        }
        if ($("#defaultTabName").val() === "cust") {
            $("#filterQtrDropdown").addClass("hidden");
            $("#customLink").click();
        }
        if ($("#webEnabled").val() === "false") {
            $("#webinarsLink").addClass("hidden");
        }
        if ($("#vidEnabled").val() === "false") {
            $("#videosLink").addClass("hidden");
        }
        if ($("#custEnabled").val() === "false") {
            $("#customLink").addClass("hidden");
        }
        if ($("#defaultTabName").val() === "web") {
            $("#webinarsHeader").removeClass("hidden");
            $("#webinars").removeClass("hidden");
            $("#webinarsNav").addClass("active");
        }
        if ($("#defaultTabName").val() === "vid") {
            $("#videosHeader").removeClass("hidden");
            $("#videos").removeClass("hidden");
            $("#videosNav").addClass("active");
        }
        if ($("#defaultTabName").val() === "cust") {
            $("#customHeader").removeClass("hidden");
            $("#custom").removeClass("hidden");
            $("#customNav").addClass("active");
        }
        if ($("#showFilterList").val() === "false") {
            $("#filterDropdown").addClass("hidden");
        }
        if ($("#showClearButton").val() === "false") {
            $("#clearFilters").addClass("hidden");
        }
        $(".next").on("click", function() {
            const currentPanel: any = $(this).closest("div");
            const nextPanel: any = currentPanel.find("div");
            if (nextPanel.attr("class").indexOf("hidden") !== 1) {
                nextPanel.removeClass("hidden");
                $(this).css("display", "none");
                $(this).parent().css("display", "none");
            }
            return false;
        });
        $("#dropdown-menu").empty();
        let strFilterList: any = "";
        strFilterList = $("#appFilters").val();
        // tslint:disable-next-line:only-arrow-functions
        $.each(strFilterList.split(";").reverse(), function(index, item) {
            $("#dropdown-menu").prepend("<a class='dropdown-item' href='#' id='fil" + item.replace(" ", "") + "'>" + item + "</a>");
            // tslint:disable-next-line:only-arrow-functions
            $("#fil" + item.replace(" ", "")).bind("click", function() {
                $("[class='col-sm-4 d-flex pb-3']").attr("style", "display: none !important");
                $("h5[class='card-title']:contains('" + item + "')").parent().parent().parent().attr("style", "display: '' !important");
                $("h6[class='card-subtitle']:contains('" + item + "')").parent().parent().parent().attr("style", "display: '' !important");
                $("span[class='app-text']:contains('" + item + "')").parent().parent().parent().parent().attr("style", "display: '' !important");
            });
        });
        $("#dropdown-menu").append("<div class='dropdown-divider'></div>");
        $("#dropdown-menu").append("<a class='dropdown-item' href='#' id='filClearAll'>Clear Filters</a>");
        $("#filterForm").empty();
        let strQtrList: any = "";
        strQtrList = $("#qtrFilters").val();
        if (strQtrList !== "") {
            // tslint:disable-next-line:only-arrow-functions
            $.each(strQtrList.split(";").reverse(), function(qtrIndex, item) {
                $("#filterForm").prepend("<div class='form-check'>&nbsp;&nbsp;&nbsp;<input type='checkbox' id='chk" + item.replace(" ", "") + "'></input>&nbsp;&nbsp;<label>" + item.replace(" ", "") + "</label></div>");
                // tslint:disable-next-line:only-arrow-functions
                $("#chk" + item.replace(" ", "")).bind("click", function() {
                    $("[class='col-sm-4 d-flex pb-3']").attr("style", "display: none !important");
                    const varQtrs = strQtrList.split(";").reverse();
                    let bolAtLeastOne: boolean = false;
                    // tslint:disable-next-line:only-arrow-functions
                    $.each(varQtrs, function(index, qtrItem) {
                        if ($("#chk" + qtrItem.replace(" ", "")).prop("checked") === true) {
                            $("p[class='card-text']:contains('" + qtrItem + "')").parent().parent().parent().attr("style", "display: '' !important");
                            bolAtLeastOne = true;
                        }
                    });
                    if (bolAtLeastOne === false) {
                        $("[class^='col-sm-4 d-flex pb-3']").attr("style", "display: '' !important");
                    }
                });
            });
        }
    }
    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        let navbarHeader: any;
        let submitForm: any;
        let webinarsHeaderString: any;
        let videosHeaderString: any;
        let customHeaderString: any;
        let connCfgMade: boolean;
        let jroutput: any;
        let jroutputCfg: any;
        let arrWebinars: any;
        let arrVideos: any;
        let arrCustom: any;
        let arrCfg: any;
        let jsonurlCfg: any;
        let strEntityId: any;
        let dataCfgAvailable: boolean;
        let defaultTab: string;
        let strTZ: string;
        let dtEventDate: any;
        let dtNow: any;
        let dtCEventDate: any;
        let dtCNow: any;
        let bolCustomAssets: boolean = false;
        let bolShowFreeLink: boolean = false;
        let strNavbarBrand: string = "";
        let strHeaderLogo: string = "";
        let strCardTop: string = "";
        let strCardNew: string = "";
        let strCardThumb: string = "";
        let strExcludedApps: string = "";
        let strExcludedSKUs: string = "";
        let strNavbarAltColor: string = "";
        let strAppFilterList: string = "";
        let strShowFilterList: string = "";
        let strShowClearButton: string = "";
        let strShowReleaseList: string = "";
        let strQtrFilters: string = "";
        let strA365image: string = "";
        let strFreeText: string = "";
        let strFreeLink: string = "";
        let strFreeBtnClr: string = "";
        const webinarsBodyString: any = [];
        const videosBodyString: any = [];
        const customBodyString: any = [];
        const xmlconn = new XMLHttpRequest();
        const xmlconnCfg = new XMLHttpRequest();
        const isEnabled: any = [];
        const apiKey: any = [];
        const jsonurl: any = [];
        const intShowRec: any = [];
        const strTitle: any = [];
        const strSubtitle: any = [];
        const strJoinText: any = [];
        const isActive: any = [];
        const dataAvailable: any = [];
        const hasAPIKey: any = [];
        const orgName: any = [];
        const recIdx: any = [];
        const connMade: any = [];
        const bannerImgs: any = [];
        const currDate: Date = new Date();
        bannerImgs[0] = "Excel-color.png";
        bannerImgs[1] = "Office365-color.png";
        bannerImgs[2] = "OneDrive-color.png";
        bannerImgs[3] = "Outlook-color.png";
        bannerImgs[4] = "PowerBI-color.png";
        bannerImgs[5] = "PowerPoint-color.png";
        bannerImgs[6] = "SharePoint-color.png";
        bannerImgs[7] = "Skype-color.png";
        bannerImgs[8] = "Teams-color.png";
        bannerImgs[9] = "Windows-color.png";
        bannerImgs[10] = "Word-color.png";
        bannerImgs[11] = "Yammer-color.png";
        strA365image = "https://a365.vitalyst.com/images/A365%20Logo.png";
        strEntityId = this.state.entityId;
        if (strEntityId === "This is not hosted in Microsoft Teams") {
            strEntityId = "BIGFIC-200";
        }
        defaultTab = "";
        jsonurlCfg = "https://bodachek.github.io/vitalystTeamsApp/config.json";
        dataCfgAvailable = false;
        if ("withCredentials" in xmlconnCfg) {
            try {
                xmlconnCfg.open("GET", jsonurlCfg, false);
                xmlconnCfg.send();
                connCfgMade = true;
            } catch (e) {
                connCfgMade = false;
                dataCfgAvailable = false;
            }
        }
        if (xmlconnCfg.status === 200 && typeof(strEntityId) !== "undefined") {
            jroutputCfg = JSON.parse(xmlconnCfg.responseText);
            arrCfg = Object.keys(jroutputCfg).map((key) => [key, jroutputCfg[key]]);
            // tslint:disable-next-line:prefer-for-of
            for (let intkey = 0; intkey < arrCfg[0][1].length; intkey++) {
                if (arrCfg[0][1][intkey].account === strEntityId) {
                    isEnabled[0] = arrCfg[0][1][intkey].webConfig[0].enabled;
                    apiKey[0] = arrCfg[0][1][intkey].webConfig[0].apiKey;
                    jsonurl[0] = arrCfg[0][1][intkey].webConfig[0].jsonurl;
                    intShowRec[0] = arrCfg[0][1][intkey].webConfig[0].intShowRec;
                    strTitle[0] = arrCfg[0][1][intkey].webConfig[0].strTitle;
                    strSubtitle[0] = arrCfg[0][1][intkey].webConfig[0].strSubtitle;
                    strJoinText[0] = arrCfg[0][1][intkey].webConfig[0].strJoinText;
                    isEnabled[1] = arrCfg[0][1][intkey].vidConfig[0].enabled;
                    apiKey[1] = arrCfg[0][1][intkey].vidConfig[0].apiKey;
                    jsonurl[1] = arrCfg[0][1][intkey].vidConfig[0].jsonurl;
                    intShowRec[1] = arrCfg[0][1][intkey].vidConfig[0].intShowRec;
                    strTitle[1] = arrCfg[0][1][intkey].vidConfig[0].strTitle;
                    strSubtitle[1] = arrCfg[0][1][intkey].vidConfig[0].strSubtitle;
                    strJoinText[1] = arrCfg[0][1][intkey].vidConfig[0].strJoinText;
                    isEnabled[2] = arrCfg[0][1][intkey].custConfig[0].enabled;
                    apiKey[2] = arrCfg[0][1][intkey].custConfig[0].apiKey;
                    jsonurl[2] = arrCfg[0][1][intkey].custConfig[0].jsonurl;
                    intShowRec[2] = arrCfg[0][1][intkey].custConfig[0].intShowRec;
                    strTitle[2] = arrCfg[0][1][intkey].custConfig[0].strTitle;
                    strSubtitle[2] = arrCfg[0][1][intkey].custConfig[0].strSubtitle;
                    strJoinText[2] = arrCfg[0][1][intkey].custConfig[0].strJoinText;
                    defaultTab = arrCfg[0][1][intkey].default;
                    strFreeText = arrCfg[0][1][intkey].strFreeText;
                    strFreeLink = arrCfg[0][1][intkey].strFreeLink;
                    strFreeBtnClr = arrCfg[0][1][intkey].strFreeBtnClr;
                    bolShowFreeLink = arrCfg[0][1][intkey].bolShowFreeLink;
                    bolCustomAssets = arrCfg[0][1][intkey].hasCustomAssets;
                    strExcludedApps = arrCfg[0][1][intkey].excludeApps;
                    strExcludedSKUs = arrCfg[0][1][intkey].excludeSKUs;
                    strShowFilterList = arrCfg[0][1][intkey].showFilterList;
                    strShowClearButton = arrCfg[0][1][intkey].showClearButton;
                    strShowReleaseList = arrCfg[0][1][intkey].showReleaseList;
                    if (arrCfg[0][1][intkey].navbarAltColor !== "" && arrCfg[0][1][intkey].navbarAltColor !== "undefined") {
                        strNavbarAltColor = arrCfg[0][1][intkey].navbarAltColor;
                    } else {
                        strNavbarAltColor = "#007bff";
                    }
                    if (arrCfg[0][1][intkey].appFilterList !== "" && arrCfg[0][1][intkey].appFilterList !== "undefined") {
                        strAppFilterList = arrCfg[0][1][intkey].appFilterList;
                    } else {
                        strAppFilterList = "Excel;Office 365;OneDrive;Outlook;PowerBI;PowerPoint;SharePoint;Skype;Teams;Word;Yammer";
                    }
                    if (arrCfg[0][1][intkey].qtrFilterList !== "" && arrCfg[0][1][intkey].qtrFilterList !== "undefined") {
                        strQtrFilters = arrCfg[0][1][intkey].qtrFilterList;
                    } else {
                        strQtrFilters = "Q1" + currDate.getFullYear() + ";Q2" + currDate.getFullYear() + ";Q3" + currDate.getFullYear() + ";Q4" + currDate.getFullYear();
                    }
                    dataCfgAvailable = true;
                    if (bolCustomAssets === true) {
                        strNavbarBrand = arrCfg[0][1][intkey].customAssets[0].imgNavbarBrand;
                        strHeaderLogo = arrCfg[0][1][intkey].customAssets[0].imgHeaderLogo;
                        strCardTop = arrCfg[0][1][intkey].customAssets[0].imgCardTop;
                        strCardNew = arrCfg[0][1][intkey].customAssets[0].imgCardNew;
                        strCardThumb = arrCfg[0][1][intkey].customAssets[0].imgCardThumb;
                    } else {
                        strNavbarBrand = "https://avatars.collectcdn.com/5aa6a1723256d7b631022680-5abbe87df845f444250480c6.png";
                        strHeaderLogo = "https://vitalystteamsapp.azurewebsites.net/assets/Adaptive%20Learning%20PbV.jpg";
                        strCardTop = "https://vitalystteamsapp.azurewebsites.net/assets/Powered%20by%20Vitalyst_1.jpg";
                        strCardNew = "https://bodachek.github.io/vitalystTeamsApp/new.PNG";
                        strCardThumb = "https://vitalystteamsapp.azurewebsites.net/assets/";
                    }
                    break;
                }
            }
        }

        const urlParams = new URLSearchParams(document.location.search.substring(1));

        if (typeof jsonurl[0] !== "undefined" && jsonurl[0] !== "" && dataCfgAvailable && typeof(strEntityId) !== "undefined") {
            if ("withCredentials" in xmlconn) {
                try {
                    xmlconn.open("GET", jsonurl[0], false);
                    xmlconn.send();
                    connMade[0] = true;
                } catch (e) {
                    connMade[0] = false;
                    isActive[0] = false;
                    dataAvailable[0] = false;
                }
            }
            if (xmlconn.status === 200) {
                jroutput = JSON.parse(xmlconn.responseText);
                arrWebinars = Object.keys(jroutput).map((key) => [key, jroutput[key]]);
                orgName[0] = "your organization";
                for (let intkey = 0; intkey < arrWebinars[0][1].length; intkey++) {
                    if (arrWebinars[0][1][intkey].APIKey === apiKey[0]) {
                        recIdx[0] = intkey;
                        hasAPIKey[0] = true;
                        isActive[0] = arrWebinars[0][1][recIdx[0]].active;
                        orgName[0] = arrWebinars[0][1][recIdx[0]].orgName;
                        dataAvailable[0] = true;
                        break;
                    }
                }
            } else {
                isActive[0] = false;
                orgName[0] = "your organization";
                dataAvailable[0] = false;
            }
        } else {
            isActive[0] = false;
            orgName[0] = "your organization";
            dataAvailable[0] = false;
        }

        if (typeof jsonurl[1] !== "undefined" && jsonurl[1] !== "" && dataCfgAvailable && typeof(strEntityId) !== "undefined") {
            if ("withCredentials" in xmlconn) {
                try {
                    xmlconn.open("GET", jsonurl[1], false);
                    xmlconn.send();
                    connMade[1] = true;
                } catch (e) {
                    connMade[1] = false;
                    isActive[1] = false;
                    dataAvailable[1] = false;
                }
            }
            if (xmlconn.status === 200) {
                jroutput = JSON.parse(xmlconn.responseText);
                arrVideos = Object.keys(jroutput).map((key) => [key, jroutput[key]]);
                orgName[1] = "your organization";
                for (let intkey = 0; intkey < arrVideos[0][1].length; intkey++) {
                    if (arrVideos[0][1][intkey].APIKey === apiKey[1]) {
                        recIdx[1] = intkey;
                        hasAPIKey[1] = true;
                        isActive[1] = arrVideos[0][1][recIdx[1]].active;
                        orgName[1] = arrVideos[0][1][recIdx[1]].orgName;
                        dataAvailable[1] = true;
                        break;
                    }
                }
            } else {
                isActive[1] = false;
                orgName[1] = "your organization";
                dataAvailable[1] = false;
            }
        } else {
            isActive[1] = false;
            orgName[1] = "your organization";
            dataAvailable[1] = false;
        }

        if (typeof jsonurl[2] !== "undefined" && jsonurl[2] !== "" && dataCfgAvailable && typeof(strEntityId) !== "undefined") {
            if ("withCredentials" in xmlconn) {
                try {
                    xmlconn.open("GET", jsonurl[2], false);
                    xmlconn.send();
                    connMade[2] = true;
                } catch (e) {
                    connMade[2] = false;
                    isActive[2] = false;
                    dataAvailable[2] = false;
                }
            }
            if (xmlconn.status === 200) {
                jroutput = JSON.parse(xmlconn.responseText);
                arrCustom = Object.keys(jroutput).map((key) => [key, jroutput[key]]);
                orgName[1] = "your organization";
                for (let intkey = 0; intkey < arrCustom[0][1].length; intkey++) {
                    if (arrVideos[0][1][intkey].APIKey === apiKey[2]) {
                        recIdx[2] = intkey;
                        hasAPIKey[2] = true;
                        isActive[2] = arrCustom[0][1][recIdx[2]].active;
                        orgName[2] = arrCustom[0][1][recIdx[2]].orgName;
                        dataAvailable[2] = true;
                        break;
                    }
                }
            } else {
                isActive[2] = false;
                orgName[2] = "your organization";
                dataAvailable[2] = false;
            }
        } else {
            isActive[2] = false;
            orgName[2] = "your organization";
            dataAvailable[2] = false;
        }

        const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });
        const { rem, font } = context;
        const { sizes, weights } = font;
        const styles = {
            header: { ...sizes.title, ...weights.semibold },
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall }
        };
        const altColor = {
            backgroundColor: strNavbarAltColor
        };
        navbarHeader = (
            <>
                <nav className="navbar navbar-expand navbar-expand-lg navbar-dark" style={altColor}>
                    <a className="navbar-brand" href="#">
                        <img src={strNavbarBrand} width="50px"/>
                    </a>
                    <button className="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarSupportedContent" aria-controls="navbarSupportedContent" aria-expanded="false" aria-label="Toggle navigation">
                        <span className="navbar-toggler-icon"></span>
                    </button>

                    <div className="collapse navbar-collapse" id="navbarSupportedContent">
                        <ul className="navbar-nav mr-auto">
                            <li id="webinarsNav" className="nav-item">
                                <a id="webinarsLink" className="nav-link" href="#">Webinars</a>
                            </li>
                            <li id="customNav" className="nav-item">
                                <a id="customLink" className="nav-link" href="#">Custom</a>
                            </li>
                            <li id="videosNav" className="nav-item">
                                <a id="videosLink" className="nav-link" href="#">Videos</a>
                            </li>
                            <li className="nav-item dropdown" id="filterDropdown">
                                <a className="nav-link dropdown-toggle" href="#" id="navbarDropdown" role="button" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                                    Applications
                                </a>
                                <div className="dropdown-menu" aria-labelledby="navbarDropdown" id="dropdown-menu">
                                    <div className="dropdown-divider"></div>
                                    <a className="dropdown-item" href="#" id="filClearAll">Clear Filters</a>
                                </div>
                            </li>
                            <li className="nav-item dropdown" id="filterQtrDropdown">
                                <a className="nav-link dropdown-toggle" href="#" id="navbarDropdown" role="button" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                                    New Features
                                </a>
                                <div className="dropdown-menu" aria-labelledby="navbarDropdown" id="dropdown-menu">
                                    <form id="filterForm" className="form-inline my-2 my-lg-0">
                                    </form>
                                    <a className="dropdown-item" href="#" id="filQA365">Active 365</a>
                                    <div className="dropdown-divider"></div>
                                    <a className="dropdown-item" href="#" id="filQClearAll">Clear Filters</a>
                                </div>
                            </li>
                            <li id="clearFilters" className="nav-item">
                                <a id="filClearAllNav" className="nav-link" href="#">Clear Filters</a>
                            </li>
                        </ul>
                        <form id="searchform" className="form-inline my-2 my-lg-0">
                            <input id="searchbox" className="form-control mr-sm-2" type="search" placeholder="Search..." aria-label="Search"></input>
                        </form>
                    </div>
                </nav>
            </>
        );

        submitForm = (
            <>
                <div className="modal fade" id="modalSubmit" role="dialog">
                    <div className="modal-dialog">
                        <div className="modal-content">
                            <div className="modal-header text-center">
                                <h4 className="modal-title w-100 text-primary">Confirm Submission to Eventbuilder</h4>
                            </div>
                            <div className="card w-100 bg-light mb-3" id="formSubmit">
                                <div className="card-body">
                                    <p className="card-text">
                                        <h4 id="submitTitle"></h4>
                                    </p>
                                    <p className="card-text">
                                        Event ID:
                                        <input type="text" id="submitEventId" name="submitEventId" className="form-control" required/>
                                        <input type="hidden" id="submitOccurrenceId" name="submitEventId" className="form-control" required/>
                                    </p>
                                    <p className="card-text">
                                        First Name:
                                        <input type="text" id="submitFName" name="submitFName" className="form-control" required/>
                                    </p>
                                    <p className="card-text">
                                        Last Name:
                                        <input type="text" id="submitLName" name="submitLName" className="form-control" required/>
                                    </p>
                                    <p className="card-text">
                                        E-mail Address:
                                        <input type="text" id="submitAddress" name="submitAddress" className="form-control" required/>
                                    </p>
                                    <p className="card-text text-center">
                                        <a id="submitForm" className="btn btn-primary" href="#" onClick={this.submitReg}>
                                            Submit
                                        </a>&nbsp;&nbsp;&nbsp;
                                        <a id="cancelForm" className="btn btn-primary" href="#formHeaderAnchor">
                                            Cancel
                                        </a>
                                    </p>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div className="modal fade" id="modalAppDialog" role="dialog">
                    <div className="modal-dialog">
                        <div className="modal-content">
                            <div className="card w-100" id="appInfo">
                                <div className="card-body">
                                    <h4 className="card-title"><span id="enrTitle"></span></h4>
                                    <p className="card-text">
                                        Application: <span id="enrApp"></span>
                                    </p>
                                    <p className="card-text">
                                        Description: <span id="enrDesc"></span>
                                    </p>
                                    <p id="ebembedWebinar" className="card-text">
                                    </p>
                                    <p className="card-text">
                                        <a id="closeAppDialog" className="btn btn-primary" href="#">
                                            OK
                                        </a>
                                    </p>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div className="modal fade" id="modalVideoDialog" role="dialog">
                    <div className="modal-dialog">
                        <div className="modal-content" id="dialogContainer">
                            <div className="card" id="appVideo">
                                <div className="card-body">
                                    <h4 className="card-title"><span id="videoTitle"></span></h4>
                                    <p className="card-text">
                                        Application: <span id="videoApp"></span>
                                    </p>
                                    <p className="card-text">
                                        Description: <span id="videoDesc"></span>
                                    </p>
                                    <p className="card-text">
                                        <iframe id="modalVideo" height="450" width="450" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture"></iframe>
                                    </p>
                                    <p className="card-text">
                                        <a id="closeVideoDialog" className="btn btn-primary" href="#">
                                            OK
                                        </a>
                                    </p>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </>
        );
        let strIsVisFree: string = "";
        if (bolShowFreeLink) {
            strIsVisFree = "";
        } else {
            strIsVisFree = "none";
        }
        const showFreeLink = {
            display: strIsVisFree
        };
        const buttonColor: Style = {
            "background-color": strFreeBtnClr,
            "border-color": strFreeBtnClr
        };
        if (isActive[0] && dataAvailable[0] && hasAPIKey[0] && typeof(strEntityId) !== "undefined") {
            webinarsHeaderString = (
                <>
                <div className="container-fluid">
                    <div className="row py-2">
                        <div className="card w-100" id="header">
                            <img className="card-img-top" src={strHeaderLogo} alt="Card image cap"></img>
                            <div className="card-body">
                                <h4 className="card-title">{strTitle[0]}</h4>
                                <p className="card-text">{strSubtitle[0]}</p>
                                <p className="card-text"></p>
                                <p className="card-text">
                                    <div className="text-center" style={showFreeLink}>
                                        <a id="btnFreeLink" className="btn btn-primary" style={buttonColor} href={strFreeLink} target="_blank">
                                            {strFreeText}
                                        </a>
                                    </div>
                                </p>
                            </div>
                        </div>
                    </div>
                </div>
                </>
            );
            for (let i: number = 0; i < arrWebinars[0][1][recIdx[0]].items.length; i++) {
                strTZ = Intl.DateTimeFormat().resolvedOptions().timeZone;
                dtEventDate = new Date(arrWebinars[0][1][recIdx[0]].items[i].date + " UTC").toLocaleString("en-US", {timeZone: strTZ});
                dtNow = new Date().toLocaleString("en-US", {timeZone: strTZ});
                dtCEventDate = new Date(dtEventDate);
                dtCNow = new Date(dtNow);
                let strIsVis: string = "";
                const cardStyle = {
                    "display": "block",
                    "margin": "0 auto",
                    "background": "url(" + arrWebinars[0][1][recIdx[0]].items[i].thumbImg + ") no-repeat scroll center",
                    "background-size": "cover",
                    "height": "200px",
                    "width": "100%",
                    "border-color": arrWebinars[0][1][recIdx[0]].items[i].borderColor
                };
                const innerThumbStyle = {
                    "background": "url(" + strCardThumb + bannerImgs[arrWebinars[0][1][recIdx[0]].items[i].bannerImg] + ") no-repeat center center",
                    "background-size": "61px 58px"
                };
                if (arrWebinars[0][1][recIdx[0]].items[i].isNew === true) {
                    strIsVis = "";
                } else {
                    strIsVis = "none";
                }
                const showNewLogo = {
                    display: strIsVis
                };
                const btnId = "btn" + i;
                if (dtCEventDate >= dtCNow) {
                    if (strExcludedApps.indexOf(arrWebinars[0][1][recIdx[0]].items[i].appID) === -1 && strExcludedSKUs.indexOf(arrWebinars[0][1][recIdx[0]].items[i].SKU) === -1) {
                        webinarsBodyString.push(
                        <>
                            <div className="col-sm-4 d-flex pb-3">
                                <div className="card">
                                    <div className="img-thumbnail" style={cardStyle}>
                                        <div className="innerThumb" style={innerThumbStyle}></div>
                                    </div>
                                    <div className="card-body">
                                        <h5 className="card-title">
                                            <img src={strCardNew} style={showNewLogo}></img>
                                            <br /> {arrWebinars[0][1][recIdx[0]].items[i].name}
                                        </h5>
                                        <h6 className="card-subtitle">{dtEventDate}</h6>
                                        <p />
                                        <p className="card-text">
                                            <i>{arrWebinars[0][1][recIdx[0]].items[i].description.substring(0, 100)}...</i>
                                            <a className="next" data-toggle="collapse" href="#" aria-expanded="true">
                                                (Show more)
                                            </a>
                                        </p>
                                            <div id="colSec" className="colText hidden">
                                                <p className="card-text">
                                                    <i>{arrWebinars[0][1][recIdx[0]].items[i].description}</i>
                                                </p>
                                            </div>
                                        <p />
                                        <p className="card-text">Application: <span className="app-text">{arrWebinars[0][1][recIdx[0]].items[i].application}</span></p>
                                        <p className="card-text">Instructor: {arrWebinars[0][1][recIdx[0]].items[i].author}</p>
                                        <a id={btnId} className="btn btn-primary" href="#" onClick={this.handleClick.bind(this, btnId, arrWebinars[0][1][recIdx[0]].items[i].frameURL, arrWebinars[0][1][recIdx[0]].items[i].type, arrWebinars[0][1][recIdx[0]].items[i].name, arrWebinars[0][1][recIdx[0]].items[i].application, arrWebinars[0][1][recIdx[0]].items[i].description)}>
                                            {strJoinText[0]}
                                        </a>
                                    </div>
                                </div>
                            </div>
                        </>
                        );
                    }
                }
                if (i + 1 >= intShowRec[0]) {
                    break;
                }
            }
        } else if (!isActive[0] && connMade[0] && dataAvailable[0] && hasAPIKey[0] && typeof(strEntityId) !== "undefined") {
            webinarsHeaderString = (
            <>
                <div className="container">
                    <div className="row py-2">
                        <div className="card w-100" id="header">
                            <img className="card-img-top" src={strHeaderLogo} alt="Card image cap"></img>
                            <div className="card-body">
                                <h4 className="card-title">Vitalyst Webinar Schedule Data</h4>
                                <p className="card-text">There are no Webinars scheduled for {orgName}.
                                </p>
                                <p className="card-text">
                                </p>
                            </div>
                        </div>
                    </div>
                </div>
            </>
            );
        } else if (!dataAvailable[0] && jsonurl[0] !== "" && apiKey[0] !== "" && typeof(strEntityId) !== "undefined") {
            webinarsHeaderString = (
                <>
                    <div className="container">
                        <div className="row py-2">
                            <div className="card w-100" id="header">
                                <img className="card-img-top" src={strHeaderLogo} alt="Card image cap"></img>
                                <div className="card-body">
                                    <h4 className="card-title">Vitalyst Webinar Schedule Data</h4>
                                    <p className="card-text">Please wait for a connection to the JSON data...
                                    </p>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            );
        } else {
            webinarsHeaderString = (
                <>
                    <div className="container">
                        <div className="row py-2">
                            <div className="card w-100" id="header">
                                <img className="card-img-top" src={strHeaderLogo} alt="Card image cap"></img>
                                <div className="card-body">
                                    <h4 className="card-title">Vitalyst Webinar Schedule Data</h4>
                                    <p className="card-text">There are no Webinars scheduled for {orgName}.</p>
                                    <p className="card-text">Invalid or Missing Web Part Settings.</p>
                                    <p className="card-text">A Data URL and API key are required to return information for your organization.</p>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            );
        }
        if (isActive[1] && dataAvailable[1] && hasAPIKey[1] && typeof(strEntityId) !== "undefined") {
            videosHeaderString = (
                <>
                <div className="container-fluid">
                    <div className="row py-2">
                        <div className="card w-100" id="header">
                            <img className="card-img-top" src={strHeaderLogo} alt="Card image cap"></img>
                            <div className="card-body">
                                <h4 className="card-title">{strTitle[1]}</h4>
                                <p className="card-text">{strSubtitle[1]}</p>
                                <p className="card-text"></p>
                                <p className="card-text">
                                    <div className="text-center" style={showFreeLink}>
                                        <a id="btnFreeLink" className="btn btn-primary" style={buttonColor} href={strFreeLink} target="_blank">
                                            {strFreeText}
                                        </a>
                                    </div>
                                </p>
                            </div>
                        </div>
                    </div>
                </div>
                </>
            );
            for (let i: number = 0; i < arrVideos[0][1][recIdx[1]].items.length; i++) {
                strTZ = Intl.DateTimeFormat().resolvedOptions().timeZone;
                dtEventDate = new Date(arrVideos[0][1][recIdx[1]].items[i].date + " UTC").toLocaleString("en-US", {timeZone: strTZ});
                dtNow = new Date().toLocaleString("en-US", {timeZone: strTZ});
                dtCEventDate = new Date(dtEventDate);
                dtCNow = new Date(dtNow);
                let strIsVis: string = "";
                const cardOverride = {
                    width: "100%"
                };
                const cardStyle = {
                    "display": "block",
                    "margin": "0 auto",
                    "background": "url(" + arrVideos[0][1][recIdx[1]].items[i].thumbImg + ") no-repeat scroll center",
                    "background-size": "cover",
                    "height": "200px",
                    "width": "100%",
                    "border-color": arrVideos[0][1][recIdx[1]].items[i].borderColor
                };
                let strThumbImage: string = "";
                if (arrVideos[0][1][recIdx[1]].items[i].active365 === "false") {
                    strThumbImage = strCardThumb + bannerImgs[arrVideos[0][1][recIdx[1]].items[i].bannerImg];
                } else {
                    strThumbImage = strA365image;
                }
                const innerThumbStyle = {
                    "background": "url(" + strThumbImage + ") no-repeat center center",
                    "background-size": "61px 58px"
                };
                if (arrVideos[0][1][recIdx[1]].items[i].isNew === true) {
                    strIsVis = "";
                } else {
                    strIsVis = "none";
                }
                const showNewLogo = {
                    display: strIsVis
                };
                const btnId = "btn" + i;
                if (dtCEventDate >= dtCNow) {
                    if (strExcludedApps.indexOf(arrVideos[0][1][recIdx[1]].items[i].appID) === -1 && strExcludedSKUs.indexOf(arrVideos[0][1][recIdx[1]].items[i].SKU) === -1) {
                        videosBodyString.push(
                        <>
                            <div className="col-sm-4 d-flex pb-3">
                                <div className="card" style={cardOverride}>
                                    <div className="img-thumbnail" style={cardStyle}>
                                        <div className="innerThumb" style={innerThumbStyle}></div>
                                    </div>
                                    <div className="card-body">
                                        <h5 className="card-title">
                                            <img src={strCardNew} style={showNewLogo}></img>
                                            <br /> {arrVideos[0][1][recIdx[1]].items[i].name}
                                        </h5>
                                        <h6 className="card-subtitle"></h6>
                                        <p />
                                        <p className="card-text">
                                            <i>{arrVideos[0][1][recIdx[1]].items[i].description.substring(0, 100)}...</i>
                                            <a className="next" data-toggle="collapse" href="#" aria-expanded="true">
                                                (Show more)
                                            </a>
                                        </p>
                                            <div id="colSec" className="colText hidden">
                                                <p className="card-text">
                                                    <i>{arrVideos[0][1][recIdx[1]].items[i].description}</i>
                                                </p>
                                            </div>
                                        <p />
                                        <p className="card-text">Release Date: {arrVideos[0][1][recIdx[1]].items[i].quarter}</p>
                                        <p className="card-text hidden"><span className="a365">Active365:{arrVideos[0][1][recIdx[1]].items[i].active365}</span></p>
                                        <p className="card-text">Application: <span className="app-text">{arrVideos[0][1][recIdx[1]].items[i].application}</span></p>
                                        <p className="card-text">Instructor: {arrVideos[0][1][recIdx[1]].items[i].author}</p>
                                        <a id={btnId} className="btn btn-primary"  href="#" onClick={this.handleClick.bind(this, btnId, arrVideos[0][1][recIdx[1]].items[i].frameURL, arrVideos[0][1][recIdx[1]].items[i].type, arrVideos[0][1][recIdx[1]].items[i].name, arrVideos[0][1][recIdx[1]].items[i].application, arrVideos[0][1][recIdx[1]].items[i].description)}>
                                            {strJoinText[1]}
                                        </a>
                                    </div>
                                </div>
                            </div>
                        </>
                        );
                    }
                }
                if (i + 1 >= intShowRec[1]) {
                    break;
                }
            }
        } else if (!isActive[1] && connMade[1] && dataAvailable[1] && hasAPIKey[1] && typeof(strEntityId) !== "undefined") {
            videosHeaderString = (
                <>
                    <div className="container">
                        <div className="row py-2">
                            <div className="card w-100" id="header">
                                <img className="card-img-top" src={strHeaderLogo} alt="Card image cap"></img>
                                <div className="card-body">
                                    <h4 className="card-title">Vitalyst Videos Schedule Data</h4>
                                    <p className="card-text">There are no Videos scheduled for {orgName}.
                                    </p>
                                    <p className="card-text">
                                    </p>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            );
        } else if (!dataAvailable[1] && jsonurl[1] !== "" && apiKey[1] !== "" && typeof(strEntityId) !== "undefined") {
            videosHeaderString = (
                <>
                    <div className="container">
                        <div className="row py-2">
                            <div className="card w-100" id="header">
                                <img className="card-img-top" src={strHeaderLogo} alt="Card image cap"></img>
                                <div className="card-body">
                                    <h4 className="card-title">Vitalyst Videos Schedule Data</h4>
                                    <p className="card-text">Please wait for a connection to the JSON data...
                                    </p>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            );
        } else {
            videosHeaderString = (
                <>
                    <div className="container">
                        <div className="row py-2">
                            <div className="card w-100" id="header">
                                <img className="card-img-top" src={strHeaderLogo} alt="Card image cap"></img>
                                <div className="card-body">
                                    <h4 className="card-title">Vitalyst Videos Schedule Data</h4>
                                    <p className="card-text">There are no Videos scheduled for {orgName}.</p>
                                    <p className="card-text">Invalid or Missing Web Part Settings.</p>
                                    <p className="card-text">A Data URL and API key are required to return information for your organization.</p>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            );
        }
        if (isActive[2] && dataAvailable[2] && hasAPIKey[2]) {
            customHeaderString = (
                <>
                <div className="container-fluid">
                    <div className="row py-2">
                        <div className="card w-100" id="header">
                            <img className="card-img-top" src={strHeaderLogo} alt="Card image cap"></img>
                            <div className="card-body">
                                <h4 className="card-title">{strTitle[2]}</h4>
                                <p className="card-text">{strSubtitle[2]}</p>
                                <p className="card-text"></p>
                                <p className="card-text">
                                    <div className="text-center" style={showFreeLink}>
                                        <a id="btnFreeLink" className="btn btn-primary" style={buttonColor} href={strFreeLink} target="_blank">
                                            {strFreeText}
                                        </a>
                                    </div>
                                </p>
                            </div>
                        </div>
                    </div>
                </div>
                </>
            );
            for (let i: number = 0; i < arrCustom[0][1][recIdx[2]].items.length; i++) {
                strTZ = Intl.DateTimeFormat().resolvedOptions().timeZone;
                dtEventDate = new Date(arrCustom[0][1][recIdx[2]].items[i].date + " UTC").toLocaleString("en-US", {timeZone: strTZ});
                dtNow = new Date().toLocaleString("en-US", {timeZone: strTZ});
                dtCEventDate = new Date(dtEventDate);
                dtCNow = new Date(dtNow);
                let strIsVis: string = "";
                const cardOverride = {
                    width: "100%"
                };
                const cardStyle = {
                    "display": "block",
                    "margin": "0 auto",
                    "background": "url(" + arrCustom[0][1][recIdx[2]].items[i].thumbImg + ") no-repeat scroll center",
                    "background-size": "cover",
                    "height": "200px",
                    "width": "100%",
                    "border-color": arrCustom[0][1][recIdx[2]].items[i].borderColor
                };
                const innerThumbStyle = {
                    "background": "url(" + strCardThumb + bannerImgs[arrCustom[0][1][recIdx[2]].items[i].bannerImg] + ") no-repeat center center",
                    "background-size": "61px 58px"
                };
                if (arrCustom[0][1][recIdx[2]].items[i].isNew === true) {
                    strIsVis = "";
                } else {
                    strIsVis = "none";
                }
                const showNewLogo = {
                    display: strIsVis
                };
                const btnId = "btn" + i;
                if (dtCEventDate >= dtCNow) {
                    if (strExcludedApps.indexOf(arrCustom[0][1][recIdx[2]].items[i].appID) === -1 && strExcludedSKUs.indexOf(arrCustom[0][1][recIdx[2]].items[i].SKU) === -1) {
                        customBodyString.push(
                        <>
                            <div className="col-sm-4 d-flex pb-3">
                                <div className="card" style={cardOverride}>
                                    <div className="img-thumbnail" style={cardStyle}>
                                        <div className="innerThumb" style={innerThumbStyle}></div>
                                    </div>
                                    <div className="card-body">
                                        <h5 className="card-title">
                                            <img src={strCardNew} style={showNewLogo}></img>
                                            <br /> {arrCustom[0][1][recIdx[2]].items[i].name}
                                        </h5>
                                        <h6 className="card-subtitle">{dtEventDate}</h6>
                                        <p />
                                        <p className="card-text">
                                            <i>{arrCustom[0][1][recIdx[2]].items[i].description.substring(0, 100)}...</i>
                                            <a className="next" data-toggle="collapse" href="#" aria-expanded="true">
                                                (Show more)
                                            </a>
                                        </p>
                                            <div id="colSec" className="colText hidden">
                                                <p className="card-text">
                                                    <i>{arrCustom[0][1][recIdx[2]].items[i].description}</i>
                                                </p>
                                            </div>
                                        <p />
                                        <p className="card-text">Application: <span className="app-text">{arrCustom[0][1][recIdx[2]].items[i].application}</span></p>
                                        <p className="card-text">Instructor: {arrCustom[0][1][recIdx[2]].items[i].author}</p>
                                        <a id={btnId} className="btn btn-primary"  href="#" onClick={this.handleClick.bind(this, btnId, arrCustom[0][1][recIdx[2]].items[i].frameURL, arrCustom[0][1][recIdx[2]].items[i].type, arrCustom[0][1][recIdx[2]].items[i].name, arrCustom[0][1][recIdx[2]].items[i].application, arrCustom[0][1][recIdx[2]].items[i].description)}>
                                            {strJoinText[2]}
                                        </a>
                                    </div>
                                </div>
                            </div>
                        </>
                        );
                    }
                }
                if (i + 1 >= intShowRec[2]) {
                    break;
                }
            }
        } else if (!isActive[1] && connMade[1] && dataAvailable[1] && hasAPIKey[1] && typeof(strEntityId) !== "undefined") {
            customHeaderString = (
                <>
                    <div className="container">
                        <div className="row py-2">
                            <div className="card w-100" id="header">
                                <img className="card-img-top" src={strHeaderLogo} alt="Card image cap"></img>
                                <div className="card-body">
                                    <h4 className="card-title">Vitalyst Custom Content Data</h4>
                                    <p className="card-text">There is no Custom Content scheduled for {orgName}.
                                    </p>
                                    <p className="card-text">
                                    </p>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            );
        } else if (!dataAvailable[1] && jsonurl[1] !== "" && apiKey[1] !== "" && typeof(strEntityId) !== "undefined") {
            customHeaderString = (
                <>
                    <div className="container">
                        <div className="row py-2">
                            <div className="card w-100" id="header">
                                <img className="card-img-top" src={strHeaderLogo} alt="Card image cap"></img>
                                <div className="card-body">
                                    <h4 className="card-title">Vitalyst Custom Content Data</h4>
                                    <p className="card-text">Please wait for a connection to the JSON data...
                                    </p>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            );
        } else {
            customHeaderString = (
                <>
                    <div className="container">
                        <div className="row py-2">
                            <div className="card w-100" id="header">
                                <img className="card-img-top" src={strHeaderLogo} alt="Card image cap"></img>
                                <div className="card-body">
                                    <h4 className="card-title">Vitalyst Custom Content Data</h4>
                                    <p className="card-text">There is no Custom Content for {orgName}.</p>
                                    <p className="card-text">Invalid or Missing Web Part Settings.</p>
                                    <p className="card-text">A Data URL and API key are required to return information for your organization.</p>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            );
        }

        return (
            <TeamsThemeContext.Provider value={context}>
                <Surface>
                    <Panel>
                        {submitForm}
                        <PanelHeader>
                            {navbarHeader}
                            <input type="hidden" id="defaultTabName" value={defaultTab}></input>
                            <input type="hidden" id="webEnabled" value={isEnabled[0]}></input>
                            <input type="hidden" id="vidEnabled" value={isEnabled[1]}></input>
                            <input type="hidden" id="custEnabled" value={isEnabled[2]}></input>
                            <input type="hidden" id="appFilters" value={strAppFilterList}></input>
                            <input type="hidden" id="showFilterList" value={strShowFilterList}></input>
                            <input type="hidden" id="showClearButton" value={strShowClearButton}></input>
                            <input type="hidden" id="showReleaseList" value={strShowReleaseList}></input>
                            <input type="hidden" id="qtrFilters" value={strQtrFilters}></input>
                            <div id="webinarsHeader" className="hidden">
                                {webinarsHeaderString}
                            </div>
                            <div id="videosHeader" className="hidden">
                                {videosHeaderString}
                            </div>
                            <div id="customHeader" className="hidden">
                                {customHeaderString}
                            </div>
                        </PanelHeader>
                        <PanelBody>
                            <div id="webinars" className="row equal hidden">
                                {webinarsBodyString}
                            </div>
                            <div id="videos" className="row equal hidden">
                                {videosBodyString}
                            </div>
                            <div id="custom" className="row equal hidden">
                                {customBodyString}
                            </div>
                        </PanelBody>
                        <PanelFooter>
                            <div id="bottomFooter" style={styles.footer}>
                                (C) Copyright Vitalyst, LLC.
                            </div>
                        </PanelFooter>
                    </Panel>
                </Surface>
            </TeamsThemeContext.Provider>
        );
    }
}
