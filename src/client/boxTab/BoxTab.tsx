import * as React from "react";
import {
    Provider,
    Flex,
    Text,
    Button,
    Header,
    Input,
    Dropdown,
    Accordion,
    List,
    ListItem,
    FlexItem
} from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import axios from "axios";
import * as microsoftTeams from "@microsoft/teams-js";

const qs = require("qs");

// List items
const items = [
    <ListItem
        key={"excel"}
        index={0}
        header={"Add excel extension query."}
        content={"Or type :xlsx at the end of your search."}
    ></ListItem>,
    <ListItem
        key={"docx"}
        index={1}
        header={"Add docx extension query."}
        content={"Or type :docx at the end of your search."}
    ></ListItem>,
    <ListItem
        key={"pptx"}
        index={2}
        header={"Add PPT extension query."}
        content={"Or type :pptx at the end of your search."}
    ></ListItem>,
    <ListItem
        key={"pdf"}
        index={3}
        header={"Add PDF extension query."}
        content={"Or type :pdf at the end of your search."}
    ></ListItem>
];

const ListSelectable = () => (
    <List styles={{ width: "250px" }} selectable items={items} />
);

// List of extension queries under the search bar
const AccordionPanel = () => {
    const panels = [
        {
            title: "Add an extension to your search query?",
            content: (
                <Flex key="extensions">
                    <ListSelectable />
                </Flex>
            )
        }
    ];
    return <Accordion styles={{ marginTop: "1px" }} panels={panels} />;
};
/**
 * Implementation of the Box Tab content page
 */
export const BoxTab = () => {
    const [{ inTeams, theme, context, themeString }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    //const [signedIn, setSignedIn] = useState<boolean>(false);

    useEffect(() => {
        if (inTeams === true) {
            microsoftTeams.appInitialization.notifySuccess();
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.entityId);
        }
    }, [context]);

    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex
                fill={true}
                styles={{
                    padding: "1rem 0 1rem .5rem"
                }}
                column
                hAlign={"center"}
            >
                <Flex.Item
                    styles={{
                        margin: "0 auto"
                    }}
                >
                    <div style={{ paddingLeft: 0 }}>
                        <AccordionPanel />
                    </div>
                </Flex.Item>
                {AccessTokenExists() ||
                    (RefreshTokenExists() ? location.reload() : false) || (
                        <FlexItem styles={{ margin: "10% auto" }}>
                            <Button onClick={SetCookies}>Login</Button>
                        </FlexItem>
                    )}
                <Flex.Item
                    styles={{
                        margin: "5% 0",
                        marginBottom: 0,
                        height: "100%",
                        width: "100%",
                        padding: 0,
                        minWidth: "320px"
                    }}
                >
                    <div className="container"></div>
                </Flex.Item>
            </Flex>
        </Provider>
    );
};

const GetTokenObject = async (code: string) => {
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

    return accessToken.data;
};

const GetRefreshTokenObj = async (token) => {
    const authenticationUrl = "https://api.box.com/oauth2/token";
    const clientDetails: any = await axios.get("/client");
    let accessToken = await axios.post(
        authenticationUrl,
        qs.stringify({
            client_id: `${clientDetails.data.id}`,
            client_secret: `${clientDetails.data.secret}`,
            refresh_token: `${token}`,
            grant_type: "refresh_token"
        }),
        { headers: { "Access-Control-Allow-Origin": "*" } }
    );

    return accessToken.data;
};

const GetAuthUrl = async () => {
    const url = await axios.get<string>("/auth", {
        headers: { "Access-Control-Allow-Origin": "*" }
    });
    return url.data;
};

const SetCookies = () => {
    GetAuthUrl().then((authorizationUrl) => {
        microsoftTeams.authentication.authenticate({
            url:
                "https://box-integration-tab.herokuapp.com/?url=" +
                encodeURIComponent(authorizationUrl),
            width: 600,
            height: 900,
            successCallback: function (result: string) {
                GetTokenObject(result).then(function (tokenObj) {
                    document.cookie = `access_token=${tokenObj.access_token};max-age=${tokenObj.expires_in};secure;path=/;samesite=none`;
                    document.cookie = `refresh_token=${
                        tokenObj.refresh_token
                    };max-age=${60 * 60 * 24 * 60};path=/;secure;samesite=none`; //two months
                });
                location.reload();
            },
            failureCallback: (result) => {
                console.log(result);
            }
        });
    });
};

const AccessTokenExists = (): boolean => {
    const cookieValue = document.cookie
        .split("; ")
        .find((row) => row.startsWith("access_token="))
        ?.split("=")[1];
    if (cookieValue) return true;
    return false;
};

const RefreshTokenExists = (): boolean => {
    const cookieValue = document.cookie
        .split("; ")
        .find((row) => row.startsWith("refresh_token="))
        ?.split("=")[1];
    if (cookieValue) {
        GetRefreshTokenObj(cookieValue).then((tokenObj) => {
            document.cookie = `access_token=${tokenObj.access_token};max-age=${tokenObj.expires_in};secure;path=/;samesite=none`;
            document.cookie = `refresh_token=${
                tokenObj.refresh_token
            };max-age=${60 * 60 * 24 * 60};path=/;secure;samesite=none`; //two months
        });
        return true;
    }
    return false;
};
