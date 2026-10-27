import { createHash } from "node:crypto";
import { crc32, deflateRawSync, deflateSync } from "node:zlib";
import { brand, teamsDeveloper } from "../../../connector-kit/src/branding.js";

// The Teams app package for one agent's bot (manifest + generated icons, zipped) and the
// setup script that creates the bot in the customer's tenant. docs/teams.md 5 and 10.

export const MANIFEST_VERSION = "1.24";
// Teams delivers every message where the bot is installed, not just @mentions, and the bot can
// read those conversations through Graph. Only where it was added; consented at install.
export const RSC_PERMISSIONS = ["ChannelMessage.Read.Group", "ChatMessage.Read.Chat"];

export type PackageInput = {
  displayName: string;
  teamsAppId: string;
  botAppId: string;
  manifestVersion: number;
  // The public origin and path prefix of this service, e.g. https://x.example/agent-connect.
  publicUrl: string;
};

export function manifest(input: PackageInput) {
  const name = input.displayName.trim();
  if (name.length < 1 || name.length > 30) throw new Error("A bot name is 1 to 30 characters");
  const host = new URL(input.publicUrl).host;
  const full = `${name} is an agent your organization set up in ${brand().productName}. It reads messages in the teams and chats it is added to, and answers when asked. Your administrator controls what it may do.`;
  return {
    $schema: `https://developer.microsoft.com/json-schemas/teams/v${MANIFEST_VERSION}/MicrosoftTeams.schema.json`,
    manifestVersion: MANIFEST_VERSION,
    version: `1.0.${input.manifestVersion}`,
    id: input.teamsAppId,
    // The app's developer, from the branding.
    developer: teamsDeveloper(),
    name: { short: name, full: `${name} (${brand().productName} agent)`.slice(0, 100) },
    description: { short: `${name}, an agent from ${brand().productName}`.slice(0, 80), full },
    icons: { color: "color.png", outline: "outline.png" },
    accentColor: accent(name),
    bots: [
      {
        botId: input.botAppId,
        scopes: ["personal", "team", "groupChat"],
        isNotificationOnly: false,
        supportsFiles: false,
      },
    ],
    webApplicationInfo: {
      id: input.botAppId,
      resource: `api://${host}/agent-connect/teams/${input.botAppId}`,
    },
    authorization: {
      permissions: {
        resourceSpecific: RSC_PERMISSIONS.map((name) => ({ name, type: "Application" })),
      },
    },
  };
}

export function packageZip(input: PackageInput): Buffer {
  const name = input.displayName;
  return zip([
    ["manifest.json", Buffer.from(JSON.stringify(manifest(input), null, 2))],
    ["color.png", colorIcon(name)],
    ["outline.png", outlineIcon(name)],
  ]);
}

// ---- icons: the agent's initials, from a 5x7 bitmap font, on a color from its name ----

const GLYPHS: Record<string, string> = {
  A: "01110100011000111111100011000110001",
  B: "11110100011000111110100011000111110",
  C: "01110100011000010000100001000101110",
  D: "11100100101000110001100011001011100",
  E: "11111100001000011110100001000011111",
  F: "11111100001000011110100001000010000",
  G: "01110100011000010111100011000101111",
  H: "10001100011000111111100011000110001",
  I: "01110001000010000100001000010001110",
  J: "00111000100001000010000101001001100",
  K: "10001100101010011000101001001010001",
  L: "10000100001000010000100001000011111",
  M: "10001110111010110101100011000110001",
  N: "10001100011100110101100111000110001",
  O: "01110100011000110001100011000101110",
  P: "11110100011000111110100001000010000",
  Q: "01110100011000110001101011001001101",
  R: "11110100011000111110101001001010001",
  S: "01111100001000001110000010000111110",
  T: "11111001000010000100001000010000100",
  U: "10001100011000110001100011000101110",
  V: "10001100011000110001100010101000100",
  W: "10001100011000110101101011010101010",
  X: "10001100010101000100010101000110001",
  Y: "10001100010101000100001000010000100",
  Z: "11111000010001000100010001000011111",
  "0": "01110100011001110101110011000101110",
  "1": "00100011000010000100001000010001110",
  "2": "01110100010000100010001000100011111",
  "3": "11111000100010000010000011000101110",
  "4": "00010001100101010010111110001000010",
  "5": "11111100001111000001000011000101110",
  "6": "00110010001000011110100011000101110",
  "7": "11111000010001000100010000100001000",
  "8": "01110100011000101110100011000101110",
  "9": "01110100011000101111000010001001100",
};

/** Up to two characters: the first letters of the first two words, else "AI". */
export function initials(name: string): string {
  const words = name
    .normalize("NFD")
    .replace(/\p{M}/gu, "")
    .toUpperCase()
    .split(/[^A-Z0-9]+/)
    .filter(Boolean);
  const out = words
    .slice(0, 2)
    .map((w) => w[0])
    .join("");
  return out || "AI";
}

const hue = (name: string) => createHash("sha256").update(name).digest()[0] / 255;

function hsl(h: number, s: number, l: number): [number, number, number] {
  const f = (n: number) => {
    const k = (n + h * 12) % 12;
    const a = s * Math.min(l, 1 - l);
    return Math.round(255 * (l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1))));
  };
  return [f(0), f(8), f(4)];
}

function accent(name: string) {
  return `#${hsl(hue(name), 0.55, 0.38)
    .map((v) => v.toString(16).padStart(2, "0"))
    .join("")}`.toUpperCase();
}

// Pixels of the initials, centered, `scale` pixels per font dot.
function glyphMask(text: string, size: number, scale: number): (x: number, y: number) => boolean {
  const chars = [...text].map((c) => GLYPHS[c] ?? GLYPHS.A);
  const width = (chars.length * 5 + (chars.length - 1)) * scale;
  const x0 = Math.floor((size - width) / 2);
  const y0 = Math.floor((size - 7 * scale) / 2);
  return (x, y) => {
    const gx = Math.floor((x - x0) / scale);
    const gy = Math.floor((y - y0) / scale);
    if (x < x0 || y < y0 || gy >= 7) return false;
    const index = Math.floor(gx / 6);
    const col = gx % 6;
    if (index >= chars.length || col === 5) return false;
    return chars[index][gy * 5 + col] === "1";
  };
}

/** 192×192, the initials in white on the agent's color. */
export function colorIcon(name: string): Buffer {
  const text = initials(name);
  const [r, g, b] = hsl(hue(name), 0.55, 0.38);
  const on = glyphMask(text, 192, text.length > 1 ? 12 : 16);
  return png(192, (x, y) => (on(x, y) ? [255, 255, 255, 255] : [r, g, b, 255]));
}

/** 32×32, only white and transparent pixels, as Teams requires. */
export function outlineIcon(name: string): Buffer {
  const text = initials(name);
  const on = glyphMask(text, 32, 2);
  return png(32, (x, y) => (on(x, y) ? [255, 255, 255, 255] : [0, 0, 0, 0]));
}

function chunk(type: string, data: Buffer): Buffer {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([length, body, crc]);
}

function png(size: number, pixel: (x: number, y: number) => number[]): Buffer {
  const header = Buffer.alloc(13);
  header.writeUInt32BE(size, 0);
  header.writeUInt32BE(size, 4);
  header.set([8, 6, 0, 0, 0], 8); // 8-bit RGBA
  const raw = Buffer.alloc(size * (size * 4 + 1));
  for (let y = 0; y < size; y++)
    for (let x = 0; x < size; x++) raw.set(pixel(x, y), y * (size * 4 + 1) + 1 + x * 4);
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", header),
    chunk("IDAT", deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

// ---- zip (deflate), enough for three files ----

function zip(files: [string, Buffer][]): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const [name, data] of files) {
    const packed = deflateRawSync(data);
    const nameBytes = Buffer.from(name, "utf8");
    const common = Buffer.alloc(26);
    common.writeUInt16LE(20, 0); // version needed
    common.writeUInt16LE(0, 2); // flags
    common.writeUInt16LE(8, 4); // deflate
    common.writeUInt32LE(0x00210000, 6); // 1980-01-01 00:00, for reproducible output
    common.writeUInt32LE(crc32(data), 10);
    common.writeUInt32LE(packed.length, 14);
    common.writeUInt32LE(data.length, 18);
    common.writeUInt16LE(nameBytes.length, 22);
    common.writeUInt16LE(0, 24);
    const local = Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]), common, nameBytes, packed]);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    common.copy(central, 6);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBytes);
    locals.push(local);
    offset += local.length;
  }
  const dir = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(dir.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, dir, end]);
}

// ---- setup script ----

/** The bot's app registration, by the unique name both Add to Teams and the script give it. */
const appUri = (botId: string) =>
  `https://graph.microsoft.com/v1.0/applications(uniqueName='${brand().resourcePrefix}-bot-${botId}')`;

// PowerShell reads the typographic single quotes as quotes too; each is doubled like '.
const psQuote = (s: string) => `'${s.replace(/['\u2018\u2019\u201a\u201b]/g, "$&$&")}'`;

/** One PowerShell script that creates the agent's bot in the customer's tenant: app
 *  registration with the service's certificate, Azure Bot with the Teams channel, and the
 *  Teams app in the organization's catalog. The admin pastes the printed application ID back. */
export function setupScript(input: {
  /** The bot's row id: its app registration's unique name, as Add to Teams names it. */
  botId: string;
  displayName: string;
  teamsAppId: string;
  tenantId: string;
  certificate: string;
  publicUrl: string;
}): string {
  // The bot's app id is only known once the script made the app: the package is filled in
  // there, from a manifest with a placeholder.
  const PLACEHOLDER = "00000000-0000-0000-0000-00000000b07a";
  const template = JSON.stringify(
    manifest({ ...input, botAppId: PLACEHOLDER, manifestVersion: 1 }),
    null,
    2,
  );
  const endpoint = `${input.publicUrl.replace(/\/+$/, "")}/teams/bots/{0}/messages`;
  return `# Creates the Microsoft Teams bot for the ${brand().productName} agent ${JSON.stringify(input.displayName)}.
# Needs an Azure subscription where you are Contributor (the bot uses the free tier and costs
# nothing), PowerShell 7 (Windows, macOS or Linux) with the MicrosoftTeams module
# (Install-Module MicrosoftTeams), and az login --tenant ${input.tenantId} as someone who can
# create app registrations. It signs you in to Teams itself (a Teams administrator); on macOS
# and Linux with a code you enter in a browser.
# Safe to run again: what already exists is reused.
param(
  [string]$Subscription = "",
  [string]$ResourceGroup = "${brand().teams.resourceGroup}",
  # Where the bot's Azure record is kept.
  [string]$Location = "westeurope"
)
$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $true
$Name = ${psQuote(input.displayName)}
$Tenant = ${psQuote(input.tenantId)}
$TeamsAppId = ${psQuote(input.teamsAppId)}

# 0. Check everything this needs before creating anything.
if (-not (Get-Module -ListAvailable MicrosoftTeams)) { throw "Install the Teams module first: Install-Module MicrosoftTeams" }
if ((az account show --query tenantId -o tsv) -ne $Tenant) { throw "Sign in to tenant $Tenant with az login --tenant $Tenant" }
if ($Subscription) { az account set --subscription $Subscription }
$Sub = az account show --query "{id:id,tenant:tenantId}" -o json | ConvertFrom-Json
if (-not $Sub.id -or $Sub.tenant -ne $Tenant) { throw "This needs an Azure subscription where you are Contributor. The bot uses the free tier and costs nothing. Pass one with -Subscription." }
# Windows signs in with a window; elsewhere Microsoft shows a code to enter in a browser.
try { Get-CsTenant | Out-Null } catch {
  if ($IsWindows -or $PSVersionTable.PSEdition -eq "Desktop") { Connect-MicrosoftTeams | Out-Null } else { Connect-MicrosoftTeams -UseDeviceAuthentication | Out-Null }
}
$Existing = Get-TeamsApp -ExternalId $TeamsAppId
function Retry([scriptblock]$Step) {
  for ($i = 1; ; $i++) {
    try { return & $Step } catch { if ($i -ge 5) { throw }; Start-Sleep -Seconds 10 }
  }
}
$Work = Join-Path ([IO.Path]::GetTempPath()) ("${brand().resourcePrefix}-bot-" + [guid]::NewGuid())
New-Item -ItemType Directory $Work | Out-Null

# 1. App registration in this organization only, with ${brand().productName}'s certificate (public half).
# Found again by its unique name, never by its display name, so no other app is touched.
$AppUri = ${psQuote(appUri(input.botId))}
@{ displayName = "$Name (${brand().productName})"; signInAudience = "AzureADMyOrg" } | ConvertTo-Json | Set-Content (Join-Path $Work "app.json")
az rest --method PATCH --uri $AppUri --headers "Prefer=create-if-missing" --body ("@" + (Join-Path $Work "app.json")) | Out-Null
$app = Retry { az rest --method GET --uri ($AppUri + '?$select=id,appId') | ConvertFrom-Json }
Write-Host "Application ID: $($app.appId)"
if (-not (az ad sp list --filter "appId eq '$($app.appId)'" --query "[0].id" -o tsv)) { Retry { az ad sp create --id $app.appId | Out-Null } }
Set-Content (Join-Path $Work "${brand().resourcePrefix}.pem") -NoNewline -Value @'
${input.certificate.trim()}
'@
Retry { az ad app credential reset --id $app.appId --cert ("@" + (Join-Path $Work "${brand().resourcePrefix}.pem")) --append | Out-Null }

# 2. Azure Bot (free tier) with the Teams channel, sending messages to ${brand().productName}.
$Bot = "${brand().resourcePrefix}-" + $app.appId.Substring(0, 8)
az provider register -n Microsoft.BotService --wait
az group create -n $ResourceGroup -l $Location | Out-Null
$found = az resource list -g $ResourceGroup --resource-type Microsoft.BotService/botServices --query "[?name=='$Bot'].id" -o tsv
# The name holds only 8 hex of the app id, so a bot made for another app may hold it.
if ($found -and (az resource show --ids $found --query properties.msaAppId -o tsv) -ne $app.appId) { throw "An Azure Bot named $Bot in $ResourceGroup belongs to another app, so it is left as it is. Rename or remove it, then run this script again." }
if (-not $found) { Retry { az bot create -g $ResourceGroup -n $Bot --app-type SingleTenant --appid $app.appId --tenant-id $Tenant --endpoint (${psQuote(endpoint)} -f $app.appId) --sku F0 --display-name $Name | Out-Null } }
try { az bot msteams show -g $ResourceGroup -n $Bot 2>$null | Out-Null } catch { az bot msteams create -g $ResourceGroup -n $Bot | Out-Null }

# 3. The Teams app, published to your organization's app catalog.
$Manifest = @'
${template}
'@
Set-Content (Join-Path $Work "manifest.json") -Value $Manifest.Replace(${psQuote(PLACEHOLDER)}, $app.appId)
[IO.File]::WriteAllBytes((Join-Path $Work "color.png"), [Convert]::FromBase64String(${psQuote(colorIcon(input.displayName).toString("base64"))}))
[IO.File]::WriteAllBytes((Join-Path $Work "outline.png"), [Convert]::FromBase64String(${psQuote(outlineIcon(input.displayName).toString("base64"))}))
$Zip = Join-Path $Work "agent.zip"
Compress-Archive -Path (Join-Path $Work "manifest.json"), (Join-Path $Work "color.png"), (Join-Path $Work "outline.png") -DestinationPath $Zip
if ($Existing) { Set-TeamsApp -Id $Existing.Id -Path $Zip | Out-Null } else { New-TeamsApp -DistributionMethod organization -Path $Zip | Out-Null }

Write-Host ""
Write-Host "Done. Paste this application ID into ${brand().productName} before you add the bot to a team: $($app.appId)"
`;
}

/** What removes the bot from the customer's tenant again. `catalogOnly`: the service already
 *  deleted the rest, so only the Teams app (which needs a signed-in administrator) is left. */
export function removalScript(input: {
  botId: string;
  botAppId: string | null;
  teamsAppId: string;
  catalogOnly?: boolean;
}): string {
  const lines = [
    "# Removes the agent's Teams bot from your organization. Run in PowerShell 7.",
    "# Or, without PowerShell: Teams admin center, Teams apps, Manage apps, open the app, Delete.",
    'if ($IsWindows -or $PSVersionTable.PSEdition -eq "Desktop") { Connect-MicrosoftTeams } else { Connect-MicrosoftTeams -UseDeviceAuthentication }',
    `Get-TeamsApp -ExternalId ${psQuote(input.teamsAppId)} | ForEach-Object { Remove-TeamsApp -Id $_.Id }`,
  ];
  if (input.botAppId && !input.catalogOnly) {
    // Only an app registration carrying this bot's unique name is the service's to delete.
    const unique = `${brand().resourcePrefix}-bot-${input.botId}`;
    lines.push(
      `$Made = az rest --method GET --uri ${psQuote(`https://graph.microsoft.com/v1.0/applications(appId='${input.botAppId}')?$select=uniqueName`)} --query uniqueName -o tsv`,
      `if ($Made -eq ${psQuote(unique)}) {`,
      `  az resource list --resource-type Microsoft.BotService/botServices --query "[?name=='${brand().resourcePrefix}-${input.botAppId.slice(0, 8)}'].id" -o tsv | ForEach-Object { if ((az resource show --ids $_ --query properties.msaAppId -o tsv) -eq ${psQuote(input.botAppId)}) { az resource delete --ids $_ } }`,
      `  az ad app delete --id ${input.botAppId}`,
      `} else { Write-Host "App registration ${input.botAppId} was not made by this bot's setup, so it is left as it is. Delete it and its Azure Bot yourself if they are the bot's." }`,
    );
  }
  return lines.join("\n") + "\n";
}
