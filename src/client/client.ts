// Default entry point for client scripts
// Automatically generated
// Please avoid from modifying to much...
import * as ReactDOM from "react-dom";
import * as React from "react";
import axios from "axios";
import * as microsoftTeams from "@microsoft/teams-js";

const qs = require("qs");
export const render = (type: any, element: HTMLElement) => {
    ReactDOM.render(React.createElement(type, {}), element);
};

export const getToken = async function (code: string) {
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

// export const getContentUrl = () => {
//     const context: any = microsoftTeams.getContext(
//         (context: microsoftTeams.Context) => {
//             return context;
//         }
//     );
//     return context.frameContext.contentUrl;
// };

// Automatically added for the boxTab tab
export * from "./boxTab/BoxTab";
export * from "./boxTab/BoxTabConfig";
export * from "./boxTab/BoxTabRemove";
