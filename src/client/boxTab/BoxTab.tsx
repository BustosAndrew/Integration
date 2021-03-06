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
import ls from "localstorage-slim";
// ls.config.encrypt = true;

const qs = require("qs");

// List items
const items = [
    <ListItem
        key={"excel"}
        index={0}
        header={"Add excel extension query."}
        content={"Or type .xlsx at the end of your search."}
        onClick={() => {
            AppendSearchQuery(".xlsx");
        }}
    ></ListItem>,
    <ListItem
        key={"docx"}
        index={1}
        header={"Add docx extension query."}
        content={"Or type .docx at the end of your search."}
        onClick={() => {
            AppendSearchQuery(".docx");
        }}
    ></ListItem>,
    <ListItem
        key={"pptx"}
        index={2}
        header={"Add PPT extension query."}
        content={"Or type .pptx at the end of your search."}
        onClick={() => {
            AppendSearchQuery(".pptx");
        }}
    ></ListItem>,
    <ListItem
        key={"pdf"}
        index={3}
        header={"Add PDF extension query."}
        content={"Or type .pdf at the end of your search."}
        onClick={() => {
            AppendSearchQuery(".pdf");
        }}
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
    const [showLogin, setShowLogin] = useState<boolean>(false);
    const [tokenObj, setTokenObj] = useState<any>();

    useEffect(() => {
        if (inTeams === true) {
            console.log(tokenObj);
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

    useEffect(() => {
        if (AccessTokenExists()) {
            console.log(tokenObj);
            setShowLogin(false);
        } else if (RefreshTokenExists()) {
            setShowLogin(false);
            GetRefreshTokenObj(ls.get("refresh_token")).then((data) => {
                ls.set("access_token", `${data.access_token}`, { ttl: 3600 });
                ls.set("refresh_token", `${data.refresh_token}`, {
                    ttl: 3600 * 24 * 60
                });
                setTokenObj(data);
                location.reload();
            });
        } else {
            setShowLogin(true);
        }
    }, []);

    const SetCookies = () => {
        GetAuthUrl().then((authorizationUrl) => {
            microsoftTeams.authentication.authenticate({
                url:
                    "https://box-integration-tab.herokuapp.com/?url=" +
                    encodeURIComponent(authorizationUrl),
                width: 600,
                height: 900,
                successCallback: function (result: string) {
                    GetTokenObject(result).then(function (data) {
                        ls.set("access_token", `${data.access_token}`, {
                            ttl: 3600
                        });
                        ls.set("refresh_token", `${data.refresh_token}`, {
                            ttl: 3600 * 24 * 60
                        });
                        setTokenObj(data);
                        location.reload();
                    });
                },
                failureCallback: (result) => {
                    setShowLogin(true);
                    ls.set("failed", "failed", { ttl: 30 });
                    // location.reload();
                }
            });
        });
    };

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
                <Flex.Item styles={{ margin: "auto" }}>
                    <AccordionPanel />
                </Flex.Item>
                {showLogin ? (
                    <FlexItem styles={{ margin: "10% auto" }}>
                        <Button onClick={SetCookies}>Login</Button>
                    </FlexItem>
                ) : null}
                <Flex.Item
                    styles={{
                        margin: "5% 0",
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
    const params = new URLSearchParams({
        grant_type: "authorization_code",
        code: code,
        client_id: `${clientDetails.data.id}`,
        client_secret: `${clientDetails.data.secret}`
    });

    const tokenObj = await axios.post(authenticationUrl, params, {
        headers: { "Access-Control-Allow-Origin": "*" }
    });
    // let accessToken = await axios.post(
    //     authenticationUrl,
    //     qs.stringify({
    //         grant_type: "authorization_code",
    //         code: code,
    //         client_id: `${clientDetails.data.id}`,
    //         client_secret: `${clientDetails.data.secret}`
    //     }),
    //     { headers: { "Access-Control-Allow-Origin": "*" } }
    // );

    return tokenObj.data;
};

const GetRefreshTokenObj = async (token) => {
    const authenticationUrl = "https://api.box.com/oauth2/token";
    const clientDetails: any = await axios.get("/client");
    let accessToken = await axios.post(
        authenticationUrl,
        qs.stringify({
            client_id: `${clientDetails.data.id}`,
            client_secret: `${clientDetails.data.secret}`,
            refresh_token: token,
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

const AccessTokenExists = (): boolean => {
    const cookieValue = ls.get("access_token");
    if (cookieValue) return true;
    return false;
};

const RefreshTokenExists = (): boolean => {
    const cookieValue = ls.get("refresh_token");
    if (cookieValue) {
        return true;
    }
    return false;
};

const AppendSearchQuery = (extenstion: string): void => {
    const searchBar: any = document.querySelector(".be-search")?.children;
    searchBar[0].value += " " + extenstion;
};
