import assert from "node:assert/strict";
import { test } from "node:test";
import { connectorSchema } from "../../packages/connector-kit/src/connector.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";
import { draftFromMetadata } from "../../scripts/odata-to-connector.ts";
import odataModule from "./index.ts";

// scripts/odata-to-connector.ts: $metadata of both versions drafted into connectors that load.

const V2 = `<?xml version="1.0" encoding="utf-8"?>
<edmx:Edmx Version="1.0" xmlns:edmx="http://schemas.microsoft.com/ado/2007/06/edmx" xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata">
  <edmx:DataServices m:DataServiceVersion="2.0">
    <Schema Namespace="API_SALES_ORDER_SRV" xmlns="http://schemas.microsoft.com/ado/2008/09/edm">
      <EntityType Name="A_SalesOrderType">
        <Key><PropertyRef Name="SalesOrder"/></Key>
        <Property Name="SalesOrder" Type="Edm.String" Nullable="false" MaxLength="10"/>
        <Property Name="SalesOrganization" Type="Edm.String" MaxLength="4"/>
        <Property Name="CreationDate" Type="Edm.DateTime"/>
        <Property Name="TotalNetAmount" Type="Edm.Decimal" Precision="16" Scale="3"/>
        <NavigationProperty Name="to_Item" Relationship="x" FromRole="a" ToRole="b"/>
      </EntityType>
      <EntityType Name="BinaryType"><Key><PropertyRef Name="Blob"/></Key><Property Name="Blob" Type="Edm.Binary"/></EntityType>
      <EntityContainer Name="API_SALES_ORDER_SRV_Entities" m:IsDefaultEntityContainer="true">
        <EntitySet Name="A_SalesOrder" EntityType="API_SALES_ORDER_SRV.A_SalesOrderType"/>
        <EntitySet Name="Binaries" EntityType="API_SALES_ORDER_SRV.BinaryType"/>
        <FunctionImport Name="ReleaseApprovalRequest" ReturnType="API_SALES_ORDER_SRV.A_SalesOrderType" m:HttpMethod="POST">
          <Parameter Name="SalesOrder" Type="Edm.String" Mode="In"/>
        </FunctionImport>
      </EntityContainer>
    </Schema>
  </edmx:DataServices>
</edmx:Edmx>`;

const V4 = `<?xml version="1.0" encoding="utf-8"?>
<edmx:Edmx Version="4.0" xmlns:edmx="http://docs.oasis-open.org/odata/ns/edmx">
  <edmx:DataServices>
    <Schema Namespace="Microsoft.Dynamics.CRM" xmlns="http://docs.oasis-open.org/odata/ns/edm">
      <EntityType Name="account">
        <Key><PropertyRef Name="accountid"/></Key>
        <Property Name="accountid" Type="Edm.Guid"/>
        <Property Name="name" Type="Edm.String"/>
        <Property Name="modifiedon" Type="Edm.DateTimeOffset"/>
      </EntityType>
      <Function Name="WhoAmI"><ReturnType Type="mscrm.WhoAmIResponse"/></Function>
      <Action Name="QualifyLead"><Parameter Name="LeadId" Type="Edm.Guid"/><Parameter Name="Options" Type="mscrm.Options"/></Action>
      <EntityContainer Name="System">
        <EntitySet Name="accounts" EntityType="Microsoft.Dynamics.CRM.account"/>
        <FunctionImport Name="WhoAmI" Function="Microsoft.Dynamics.CRM.WhoAmI"/>
        <ActionImport Name="QualifyLead" Action="Microsoft.Dynamics.CRM.QualifyLead"/>
      </EntityContainer>
    </Schema>
  </edmx:DataServices>
</edmx:Edmx>`;

const load = (connector: Record<string, unknown>) => {
  const parsed = connectorSchema.parse(connector);
  assertToolsImplemented([
    { connector: parsed, ...odataModule.create({ connector: parsed } as never) },
  ]);
  return parsed.odata!;
};

test("SAP $metadata (v2) drafts list, get and its function import", () => {
  const { connector, notes } = draftFromMetadata(V2, {
    root: "https://s4.example.com/sap/opu/odata/sap/API_SALES_ORDER_SRV",
    id: "sap-orders",
    prefix: "sap",
  });
  const od = load(connector);
  assert.equal(od.version, "2");
  assert.equal(od.csrf, true);
  assert.deepEqual(
    od.tools.map((t) => [t.name, t.kind]),
    [
      ["sap_list_a_sales_order", "list"],
      ["sap_get_a_sales_order", "get"],
      ["sap_release_approval_request", "action"],
    ],
  );
  const [list, get] = od.tools;
  assert.deepEqual(list.select, [
    "SalesOrder",
    "SalesOrganization",
    "CreationDate",
    "TotalNetAmount",
  ]);
  assert.deepEqual(get.key, {
    sales_order: { type: "string", description: "SalesOrder", property: "SalesOrder" },
  });
  assert.ok(notes.some((n) => /Binaries: skipped/.test(n)));
  assert.deepEqual(
    (connector.capabilities as { id: string }[]).map((c) => c.id),
    ["read", "actions"],
  );
});

test("Dynamics $metadata (v4) drafts entity sets, function and action imports", () => {
  const { connector, notes } = draftFromMetadata(V4, {
    root: "https://org.crm.dynamics.com/api/data/v9.2",
    prefix: "crm",
    sets: ["accounts"],
  });
  const od = load(connector);
  assert.equal(od.version, "4");
  assert.equal(od.csrf, false);
  const names = od.tools.map((t) => t.name);
  assert.deepEqual(names, [
    "crm_list_accounts",
    "crm_get_accounts",
    "crm_who_am_i",
    "crm_qualify_lead",
  ]);
  assert.equal(od.tools[1].key.accountid.type, "guid");
  assert.deepEqual(od.tools[3].params, {
    lead_id: { type: "guid", description: "LeadId", property: "LeadId" },
  });
  assert.ok(notes.some((n) => /Options has a type to map by hand/.test(n)));
});
