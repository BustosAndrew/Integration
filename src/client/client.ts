// Default entry point for client scripts
// Automatically generated
// Please avoid from modifying to much...
import * as ReactDOM from "react-dom";
import * as React from "react";
import axios, { AxiosInstance } from "axios";
import * as microsoftTeams from "@microsoft/teams-js";

const qs = require("qs");
export const render = (type: any, element: HTMLElement) => {
    ReactDOM.render(React.createElement(type, {}), element);
};

export const getToken = async function (code: string | undefined) {
    const authenticationUrl = "https://api.box.com/oauth2/token";
    const clientDetails: any = await axios.get("/client");
    let accessToken = await axios.post(
        authenticationUrl,
        qs.stringify({
            grant_type: "authorization_code",
            code: code,
            client_id: `${clientDetails.data.id}`,
            client_secret: `${clientDetails.data.secret}`
        }),
        { headers: { "Access-Control-Allow-Origin": "*" } }
    );

    return accessToken.data.access_token;
};

export const getAuthUrl = async () => {
    const url = await axios.get<string>("/auth", {
        headers: { "Access-Control-Allow-Origin": "*" }
    });
    return url.data;
};

export const getParentUrl = () => {
    let isInIframe: boolean = parent !== window,
        parentUrl: string = "";

    if (isInIframe) parentUrl = document.referrer;
    if (parentUrl !== "") console.log(parentUrl);
    return parentUrl;
};

// export const setAuthPopup = (authUrl : string) => {
// microsoftTeams.authentication.authenticate({
//                         url: authUrl,
//                         width: 600,
//                         height: 535,
//                         successCallback: function (result) {
//                             getToken(result)
//                                 .then(function (access_token) {
//                                     let folderId = "0";
//                                     const contentExplorer = new Box.ContentExplorer();
//                                     contentExplorer.show(
//                                         folderId,
//                                         `${access_token}`,
//                                         {
//                                             container: ".container"
//                                         }
//                                     );
//                                 })
// }

// Automatically added for the boxTab tab
export * from "./boxTab/BoxTab";
export * from "./boxTab/BoxTabConfig";
export * from "./boxTab/BoxTabRemove";
