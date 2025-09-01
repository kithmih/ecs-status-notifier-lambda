import { IncomingWebhook } from "ms-teams-webhook";
import { URL } from "url";

// --- Configuration ---
const teamsWebhookUrl =
  process.env.TEAMS_WEBHOOK_URL || "YOUR_WEBHOOK_URL_HERE";

// --- Main Lambda Handler ---

/**
 * The main handler for the AWS Lambda function.
 *
 * @param {object} event The event object from Amazon EventBridge.
 * @returns {Promise<object>} The response object indicating the status of the function execution.
 */
export const handler = async (event) => {
  console.log("Received event:", JSON.stringify(event, null, 2));

  if (!isValidWebhookUrl(teamsWebhookUrl)) {
    const errorMessage = "Invalid or missing Microsoft Teams Webhook URL.";
    console.error(errorMessage);
    return {
      statusCode: 400,
      body: JSON.stringify({ message: errorMessage }),
    };
  }

  const webhook = new IncomingWebhook(teamsWebhookUrl);

  try {
    const card = createAdaptiveCard(event);
    await webhook.send(card);
    console.log("Successfully sent notification to Microsoft Teams.");
    return {
      statusCode: 200,
      body: JSON.stringify({ message: "Notification sent successfully!" }),
    };
  } catch (error) {
    console.error("Error sending notification to Teams:", error);
    return {
      statusCode: 500,
      body: JSON.stringify({
        message: "Failed to send notification.",
        error: error.message,
      }),
    };
  }
};

// --- Helper Functions ---

/**
 * Creates an Adaptive Card for Microsoft Teams based on the ECS event.
 * @param {object} event The event object from Amazon EventBridge.
 * @returns {object} The formatted Adaptive Card payload.
 */
function createAdaptiveCard(event) {
  const { detail, region, time, account } = event;
  const {
    desiredStatus,
    lastStatus,
    taskDefinitionArn,
    clusterArn,
    containers,
    group,
    stoppedReason
  } = detail;

  const taskDefinitionName = taskDefinitionArn.split("/").pop();
  const clusterName = clusterArn.split("/").pop();
  const containerReason = stoppedReason ?? "N/A";
  const eventDate = new Date(time);
  const istTimeString = eventDate.toLocaleString('en-IN', {
    timeZone: 'Asia/Kolkata',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: true,
  });
  const serviceName = group.split(":")[1];

  // Determine color and status message based on the event
  const { color, statusMessage, emoji } = getStatusDetails(
    lastStatus,
    desiredStatus
  );

  const card = {
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    themeColor: color,
    summary: `ECS Service Notification: ${clusterName}`,
    sections: [
      {
        activityTitle: `${emoji} **ECS Task Status Change: ${statusMessage}**`,
        activitySubtitle: `On ${istTimeString}`,
        facts: [
          { name: "Cluster:", value: clusterName },
          { name: "Service:", value: serviceName },
          { name: "Task Definition:", value: taskDefinitionName },
          { name: "Last Status:", value: `**${lastStatus}**` },
          { name: "Desired Status:", value: `**${desiredStatus}**` },
          { name: "Reason:", value: containerReason },
          { name: "AWS Account:", value: account },
          { name: "Region:", value: region },
        ],
        markdown: true,
      },
    ],
    potentialAction: [
      {
        "@type": "OpenUri",
        name: "View ECS Cluster in Console",
        targets: [
          {
            os: "default",
            uri: `https://${region}.console.aws.amazon.com/ecs/v2/clusters/${clusterName}/services?region=${region}`,
          },
        ],
      },
    ],
  };

  return card;
}

// Determines the message based on task status.
function getStatusDetails(lastStatus, desiredStatus) {
  let color = "0078D4"; // Blue
  let statusMessage = `${lastStatus}`;
  let emoji = "ℹ️";

  if (lastStatus === "STOPPED") {
    color = "FF0000"; // Red
    statusMessage = "Task Stopped";
    emoji = "❌";
  } else if (lastStatus === "RUNNING" && desiredStatus === "RUNNING") {
    color = "28A745"; // Green
    statusMessage = "Task is Running";
    emoji = "✅";
  } else if (lastStatus === "PROVISIONING" || lastStatus === "PENDING") {
    color = "FFC107"; // Yellow
    statusMessage = "Task is Starting";
    emoji = "⏳";
  }

  return { color, statusMessage, emoji };
}

// Validates if the provided string is a valid URL.
function isValidWebhookUrl(urlString) {
  if (
    !urlString ||
    typeof urlString !== "string" ||
    !urlString.startsWith("https://")
  ) {
    return false;
  }
  try {
    new URL(urlString);
    return true;
  } catch (e) {
    return false;
  }
}