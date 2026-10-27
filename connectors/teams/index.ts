import { microsoftApp } from "../microsoft/client.js";
import { teamsTools } from "./tools.js";

// The connection only holds the ceiling; calls go out as each agent's own bot.
export default microsoftApp("teams", teamsTools);
