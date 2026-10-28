import dellTechdirect from "../../../connectors/dell-techdirect/index.js";
import dynamics365 from "../../../connectors/dynamics-365/index.js";
import entra from "../../../connectors/entra/index.js";
import fedex from "../../../connectors/fedex/index.js";
import gong from "../../../connectors/gong/index.js";
import intune from "../../../connectors/intune/index.js";
import jamf from "../../../connectors/jamf/index.js";
import jira from "../../../connectors/jira/index.js";
import netsuite from "../../../connectors/netsuite/index.js";
import okta from "../../../connectors/okta/index.js";
import outlookCalendar from "../../../connectors/outlook-calendar/index.js";
import outlookMail from "../../../connectors/outlook-mail/index.js";
import salesforce from "../../../connectors/salesforce/index.js";
import sapS4hana from "../../../connectors/sap-s4hana/index.js";
import servicenow from "../../../connectors/servicenow/index.js";
import sharepointFiles from "../../../connectors/sharepoint-files/index.js";
import slack from "../../../connectors/slack/index.js";
import teams from "../../../connectors/teams/index.js";
import workday from "../../../connectors/workday/index.js";
import type { ConnectorModule } from "../../connector-kit/src/connector.js";

export type { ConnectorModule };

/** The code behind each connector folder in connectors/ (index.ts beside its connector.json). */
export const builtinConnectors: ConnectorModule[] = [
  jira,
  outlookMail,
  outlookCalendar,
  sharepointFiles,
  salesforce,
  servicenow,
  fedex,
  dellTechdirect,
  teams,
  entra,
  intune,
  slack,
  okta,
  jamf,
  sapS4hana,
  workday,
  netsuite,
  dynamics365,
  gong,
];
