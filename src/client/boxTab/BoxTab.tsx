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
    ListItem
} from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

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

const inputItems = ["pdf", "docx", "pptx", "xlsx"];
// Filter button to the right of the search bar
const FilterSearchMultiple = () => (
    <Dropdown
        fluid
        search
        multiple
        items={inputItems}
        placeholder="Filter"
        noResultsMessage="N/A"
    />
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
    return (
        <Accordion
            styles={{ marginTop: "1px" }}
            defaultActiveIndex={[0]}
            panels={panels}
        />
    );
};
/**
 * Implementation of the Box Tab content page
 */
export const BoxTab = () => {
    const [{ inTeams, theme, context, themeString }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();

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
                    padding: ".8rem 0 .8rem .5rem"
                }}
                column
            >
                <Flex
                    fill={true}
                    styles={{
                        padding: 0
                    }}
                >
                    <Flex.Item
                        styles={{ margin: "auto", marginBottom: "10px" }}
                    >
                        <div>
                            <div
                                style={{
                                    float: "left",
                                    minWidth: "250px",
                                    marginBottom: "1px"
                                }}
                            >
                                <Input fluid></Input>
                            </div>
                            <div
                                style={{
                                    float: "right",
                                    maxWidth: "100px",
                                    maxHeight: "10px"
                                }}
                            >
                                <FilterSearchMultiple />
                            </div>
                        </div>
                    </Flex.Item>
                </Flex>
                <Flex.Item styles={{ margin: "0 auto" }}>
                    <div>
                        <div style={{ minWidth: "200px" }}>
                            <AccordionPanel />
                        </div>
                    </div>
                </Flex.Item>
                <Flex.Item
                    styles={{
                        margin: "5% auto",
                        border:
                            "1px solid " +
                            (themeString === "default" ? "black" : "white"),
                        padding: "25%",
                        borderRadius: "25px"
                    }}
                >
                    <Text>List view</Text>
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
