import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/boxTab/index.html")
@PreventIframe("/boxTab/config.html")
@PreventIframe("/boxTab/remove.html")
export class BoxTab {
}
