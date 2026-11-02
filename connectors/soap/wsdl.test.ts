import assert from "node:assert/strict";
import { test } from "node:test";
import { connectorSchema } from "../../packages/connector-kit/src/connector.ts";
import { assertToolsImplemented } from "../../packages/server/src/service.ts";
import { draftFromWsdl } from "../../scripts/wsdl-to-connector.ts";
import soapModule from "./index.ts";

// scripts/wsdl-to-connector.ts: a document/literal WSDL drafted into a connector that loads.

const WSDL = `<?xml version="1.0"?>
<wsdl:definitions name="EmployeeService" targetNamespace="urn:example:hr:v1"
  xmlns:wsdl="http://schemas.xmlsoap.org/wsdl/" xmlns:soap12="http://schemas.xmlsoap.org/wsdl/soap12/"
  xmlns:tns="urn:example:hr:v1" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <wsdl:types>
    <xsd:schema targetNamespace="urn:example:hr:v1" elementFormDefault="qualified">
      <xsd:element name="GetEmployee"><xsd:complexType><xsd:sequence>
        <xsd:element name="employeeId" type="xsd:string"/>
        <xsd:element name="includeHistory" type="xsd:boolean" minOccurs="0"/>
        <xsd:element name="filter"><xsd:complexType><xsd:sequence><xsd:element name="x" type="xsd:string"/></xsd:sequence></xsd:complexType></xsd:element>
      </xsd:sequence></xsd:complexType></xsd:element>
      <xsd:element name="GetEmployeeResponse" type="tns:EmployeeResult"/>
      <xsd:complexType name="EmployeeResult"><xsd:sequence><xsd:element name="name" type="xsd:string"/></xsd:sequence></xsd:complexType>
      <xsd:element name="ListDepartments"><xsd:complexType><xsd:sequence>
        <xsd:element name="maxCount" type="xsd:int" minOccurs="0"/>
      </xsd:sequence></xsd:complexType></xsd:element>
      <xsd:element name="ListDepartmentsResponse"><xsd:complexType><xsd:sequence><xsd:element name="department" type="xsd:string" maxOccurs="unbounded"/></xsd:sequence></xsd:complexType></xsd:element>
    </xsd:schema>
  </wsdl:types>
  <wsdl:message name="GetEmployeeIn"><wsdl:part name="parameters" element="tns:GetEmployee"/></wsdl:message>
  <wsdl:message name="GetEmployeeOut"><wsdl:part name="parameters" element="tns:GetEmployeeResponse"/></wsdl:message>
  <wsdl:message name="ListIn"><wsdl:part name="parameters" element="tns:ListDepartments"/></wsdl:message>
  <wsdl:message name="ListOut"><wsdl:part name="parameters" element="tns:ListDepartmentsResponse"/></wsdl:message>
  <wsdl:portType name="EmployeePort">
    <wsdl:operation name="GetEmployee"><wsdl:input message="tns:GetEmployeeIn"/><wsdl:output message="tns:GetEmployeeOut"/></wsdl:operation>
    <wsdl:operation name="ListDepartments"><wsdl:input message="tns:ListIn"/><wsdl:output message="tns:ListOut"/></wsdl:operation>
  </wsdl:portType>
  <wsdl:binding name="EmployeeBinding" type="tns:EmployeePort">
    <soap12:binding transport="http://schemas.xmlsoap.org/soap/http"/>
    <wsdl:operation name="GetEmployee"><soap12:operation soapAction="urn:example:hr:v1/GetEmployee"/></wsdl:operation>
    <wsdl:operation name="ListDepartments"><soap12:operation soapAction="urn:example:hr:v1/ListDepartments"/></wsdl:operation>
  </wsdl:binding>
  <wsdl:service name="EmployeeService">
    <wsdl:port name="EmployeeSoap12" binding="tns:EmployeeBinding"><soap12:address location="https://hr.example.com/EmployeeService"/></wsdl:port>
  </wsdl:service>
</wsdl:definitions>`;

test("a WSDL drafts a connector that loads and runs on the soap module", () => {
  const { connector, notes } = draftFromWsdl(WSDL, { id: "hr", prefix: "hr", wsSecurity: true });
  const parsed = connectorSchema.parse(connector);
  assertToolsImplemented([
    { connector: parsed, ...soapModule.create({ connector: parsed } as never) },
  ]);
  const soap = parsed.soap!;
  assert.equal(soap.version, "1.2");
  assert.equal(soap.endpoint, "https://hr.example.com/EmployeeService");
  assert.deepEqual(soap.namespaces, { ns1: "urn:example:hr:v1" });
  const [get, list] = soap.tools;
  assert.equal(get.name, "hr_get_employee");
  assert.equal(get.soapAction, "urn:example:hr:v1/GetEmployee");
  assert.equal(
    get.body,
    "<ns1:GetEmployee><ns1:employeeId>{employee_id}</ns1:employeeId><ns1:includeHistory>{include_history}</ns1:includeHistory></ns1:GetEmployee>",
  );
  assert.deepEqual(get.params.include_history, {
    type: "boolean",
    description: "includeHistory",
    required: false,
  });
  assert.equal(get.result, "GetEmployeeResponse");
  assert.equal(list.params.max_count.type, "integer");
  // The operation without required params proves the connection.
  assert.deepEqual(soap.verify, { tool: "hr_list_departments", args: {}, show: "department" });
  assert.ok(notes.some((n) => /filter has nested fields/.test(n)));
});
