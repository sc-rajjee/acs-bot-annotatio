
require("dotenv").config();

const fs = require("fs");
const path = require("path");

const GRAPH = "https://graph.microsoft.com/v1.0";

const [meetingUrl, organizerId] = process.argv.slice(2);

const {
  GRAPH_TENANT_ID,
  GRAPH_CLIENT_ID,
  GRAPH_CLIENT_SECRET
} = process.env;

function checkConfiguration() {
  console.log("\n========== GRAPH TRANSCRIPT DIAGNOSTICS ==========");

  console.log("Tenant ID:", GRAPH_TENANT_ID || "MISSING");
  console.log("Client ID:", GRAPH_CLIENT_ID || "MISSING");
  console.log("Client Secret:", GRAPH_CLIENT_SECRET ? "CONFIGURED (hidden)" : "MISSING");
  console.log("Organizer Object ID:", organizerId || "MISSING");
  console.log("Meeting URL:", meetingUrl ? "PROVIDED (hidden)" : "MISSING");

  const missing = [];

  if (!GRAPH_TENANT_ID) missing.push("GRAPH_TENANT_ID");
  if (!GRAPH_CLIENT_ID) missing.push("GRAPH_CLIENT_ID");
  if (!GRAPH_CLIENT_SECRET) missing.push("GRAPH_CLIENT_SECRET");
  if (!organizerId) missing.push("Organizer Entra User ID");
  if (!meetingUrl) missing.push("Teams Meeting URL");

  if (missing.length) {
    throw new Error("Missing required values: " + missing.join(", "));
  }
}

async function getToken() {
  const url =
    `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    client_id: GRAPH_CLIENT_ID,
    client_secret: GRAPH_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials"
  });

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(
      `TOKEN ERROR ${response.status}: ${JSON.stringify(data.error || data)}`
    );
  }

  return data.access_token;
}

async function graphRequest(url, token, accept = "application/json") {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: accept
    }
  });

  const content = await response.text();

  if (!response.ok) {
    let error;

    try {
      error = JSON.parse(content).error;
    } catch {
      error = { message: content };
    }

    console.error("\nGRAPH API ERROR");
    console.error("HTTP Status:", response.status);
    console.error("Error Code:", error.code || "Unknown");
    console.error("Message:", error.message || content);

    throw new Error(`Graph request failed: HTTP ${response.status}`);
  }

  return { response, content };
}

function decodeClaims(token) {
  try {
    return JSON.parse(
      Buffer.from(token.split(".")[1], "base64url").toString()
    );
  } catch {
    return {};
  }
}

function parseVtt(vtt) {
  const results = {};
  const blocks = vtt.replace(/\r/g, "").split(/\n\s*\n/);

  for (const block of blocks) {
    const timing = block.match(
      /(\d{2}:\d{2}:\d{2}\.\d{3})\s+-->\s+(\d{2}:\d{2}:\d{2}\.\d{3})/
    );

    if (!timing) continue;

    const speakerMatch = block.match(/<v(?:\.[^\s>]*)?\s+([^>]+)>/i);
    const speaker = speakerMatch ? speakerMatch[1].trim() : "Unknown";

    const start = timing[1];
    const end = timing[2];

    const entry = results[speaker] || {
      speakingSegments: 0,
      timestamps: []
    };

    entry.speakingSegments++;
    entry.timestamps.push({ start, end });

    results[speaker] = entry;
  }

  return results;
}

async function main() {
  checkConfiguration();

  console.log("\n[1/4] Authenticating with Microsoft Graph...");

  const token = await getToken();

  console.log("SUCCESS: Access token acquired.");

  const claims = decodeClaims(token);

  console.log("\nTOKEN DIAGNOSTICS");
  console.log("Tenant ID:", claims.tid || "Unavailable");
  console.log("Application ID:", claims.appid || claims.azp || "Unavailable");
  console.log("Audience:", claims.aud || "Unavailable");

  const roles = claims.roles || [];

  console.log("Granted application roles:", roles.join(", ") || "NONE");

  for (const permission of [
    "OnlineMeetings.Read.All",
    "OnlineMeetingTranscript.Read.All"
  ]) {
    console.log(
      `${permission}:`,
      roles.includes(permission) ? "PRESENT" : "NOT FOUND IN TOKEN"
    );
  }

  console.log("\n[2/4] Looking up Teams meeting...");

  const filter = `JoinWebUrl eq '${meetingUrl.replace(/'/g, "''")}'`;

  const lookupUrl =
    `${GRAPH}/users/${encodeURIComponent(organizerId)}/onlineMeetings?` +
    new URLSearchParams({ "$filter": filter });

  const lookup = await graphRequest(lookupUrl, token);
  const meetingData = JSON.parse(lookup.content);

  if (!meetingData.value || meetingData.value.length === 0) {
    throw new Error(
      "No meeting found. Check organizer ID, meeting URL and meeting type."
    );
  }

  const meeting = meetingData.value[0];

  console.log("Meeting found!");
  console.log("Meeting ID:", meeting.id);
  console.log("Subject:", meeting.subject || "Not returned");
  console.log("Start:", meeting.startDateTime || "Not returned");
  console.log("End:", meeting.endDateTime || "Not returned");

  console.log("\n[3/4] Checking available transcripts...");

  const transcriptUrl =
    `${GRAPH}/users/${encodeURIComponent(organizerId)}` +
    `/onlineMeetings/${encodeURIComponent(meeting.id)}/transcripts`;

  const transcriptResponse = await graphRequest(transcriptUrl, token);
  const transcriptData = JSON.parse(transcriptResponse.content);

  const transcripts = transcriptData.value || [];

  console.log("Transcripts found:", transcripts.length);

  if (transcripts.length === 0) {
    console.log(
      "No transcript available. Confirm transcription was started and retry after Teams publishes it."
    );
    return;
  }

  console.log("\n[4/4] Downloading transcript content...");

  const outputDir = path.join(__dirname, "transcripts");

  fs.mkdirSync(outputDir, { recursive: true });

  for (const transcript of transcripts) {
    const contentUrl =
      `${GRAPH}/users/${encodeURIComponent(organizerId)}` +
      `/onlineMeetings/${encodeURIComponent(meeting.id)}` +
      `/transcripts/${encodeURIComponent(transcript.id)}/content`;

    const result = await graphRequest(contentUrl, token, "text/vtt");

    const filename =
      `meeting_${Date.now()}_${transcript.id.replace(/[^a-z0-9_-]/gi, "_")}.vtt`;

    const filePath = path.join(outputDir, filename);

    fs.writeFileSync(filePath, result.content, "utf8");

    console.log("\nTranscript saved:", filePath);

    const summary = parseVtt(result.content);

    console.log("\n========== SPEAKER SUMMARY ==========");

    for (const [speaker, data] of Object.entries(summary)) {
      console.log(`\nSpeaker: ${speaker}`);
      console.log("Speaking segments:", data.speakingSegments);

      data.timestamps.forEach((timestamp, index) => {
        console.log(
          `  ${index + 1}. ${timestamp.start} --> ${timestamp.end}`
        );
      });
    }

    if (Object.keys(summary).length === 0) {
      console.log("No speaker tags detected. Inspect the downloaded VTT.");
    }
  }

  console.log("\n========== TEST COMPLETED ==========");
}

main().catch(error => {
  console.error("\nTEST FAILED:", error.message);
  console.error("Check the diagnostics above and the Graph permission configuration.");
  process.exitCode = 1;
});
