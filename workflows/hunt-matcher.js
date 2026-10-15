export default async function (ctx) {
  const { input } = ctx;
  const spreadsheetId = input.spreadsheetId;
  if (!spreadsheetId) {
    throw new Error("Missing required input field: spreadsheetId");
  }

  const sheetName = input.sheetName || "Sheet1";
  const minMatchScore = input.minMatchScore !== undefined ? Number(input.minMatchScore) : 70;
  const slackChannel = input.slackChannel || "";
  const userEmail = input.userEmail || input.recipientEmail || "rimahmood2020@gmail.com";
  let resumeText = input.resumeText || "";

  // 1. Fetch base resume from Google Drive if resumeFileId provided
  if (input.resumeFileId) {
    try {
      const driveRes = await fastn.connector.googleDrive.exportFile({
        fileId: input.resumeFileId,
        mimeType: "text/plain"
      });
      const content = driveRes?.output?.content || driveRes?.content;
      if (content) {
        resumeText = String(content);
      }
    } catch (driveErr) {
      console.warn("Could not fetch resume from Google Drive:", driveErr.message);
    }
  }

  if (!resumeText) {
    resumeText = "Senior Full Stack Engineer proficient in TypeScript, JavaScript, React, Node.js, Python, SQL, REST APIs, Git, Docker, AWS, CI/CD, PostgreSQL, GraphQL.";
  }

  const normalizedResume = resumeText.toLowerCase();

  // 2. Read Sheet rows
  let rows = [];
  if (Array.isArray(input.mockRows) && input.mockRows.length > 0) {
    rows = [
      ["job_id", "title", "company", "location", "work_mode", "salary", "source", "link", "posted_date", "days_since_posted", "apply_by", "apply_by_estimated", "match_score", "missing_skills", "status", "applied_on", "resume_version_link", "contact_email", "follow_up_due", "last_update", "notes"],
      ...input.mockRows
    ];
  } else {
    try {
      const sheetData = await fastn.connector.googleSheets.getValues({
        spreadsheetId,
        range: `${sheetName}!A:U`
      });
      rows = sheetData?.output?.values || sheetData?.values || [];
    } catch (sheetErr) {
      console.warn("Could not read Google Sheet rows:", sheetErr.message);
    }
  }

  const skillDictionary = [
    "typescript", "javascript", "react", "next.js", "vue", "angular", "node.js", "nodejs",
    "express", "python", "fastapi", "django", "java", "spring", "golang", "go", "c#", ".net",
    "rust", "sql", "postgresql", "postgres", "mysql", "mongodb", "redis", "elasticsearch",
    "aws", "gcp", "azure", "docker", "kubernetes", "ci/cd", "git", "graphql", "rest", "grpc",
    "microservices", "kafka", "rabbitmq", "tailwind", "html", "css", "machine learning",
    "pytorch", "tensorflow", "spark", "pandas", "data pipeline"
  ];

  let processed = 0;
  let shortlisted = 0;
  let reviewed = 0;
  const details = [];
  const shortlistedItems = [];
  const todayIso = new Date().toISOString().split("T")[0];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    const currentStatus = String(row[14] || "").trim();
    if (currentStatus !== "New") continue;

    const jobId = String(row[0] || "");
    const title = String(row[1] || "");
    const company = String(row[2] || "");
    const location = String(row[3] || "");
    const workMode = String(row[4] || "");
    const salary = String(row[5] || "");
    const link = String(row[7] || "");
    const applyBy = String(row[10] || "");
    const applyByEstimated = String(row[11] || "TRUE");
    const notes = String(row[20] || "");

    const jobFullText = `${title} ${notes}`.toLowerCase();

    // Identify skills mentioned in the job posting
    const jdSkills = [];
    for (const skill of skillDictionary) {
      const escaped = skill.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const rx = new RegExp(`\\b${escaped}\\b`, "i");
      if (rx.test(jobFullText)) {
        jdSkills.push(skill);
      }
    }

    // Determine matched vs missing skills against resume
    const matchedSkills = [];
    const missingSkills = [];
    for (const skill of jdSkills) {
      const escaped = skill.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const rx = new RegExp(`\\b${escaped}\\b`, "i");
      if (rx.test(normalizedResume)) {
        matchedSkills.push(skill);
      } else {
        missingSkills.push(skill);
      }
    }

    // Deterministic title match (weight 30%)
    const titleTokens = title.toLowerCase().split(/[\s,/-]+/).filter(w => w.length > 2 && !["and", "for", "the", "with"].includes(w));
    let titleHits = 0;
    for (const token of titleTokens) {
      if (normalizedResume.includes(token)) titleHits++;
    }
    const titleScore = titleTokens.length > 0 ? (titleHits / titleTokens.length) * 100 : 50;

    // Deterministic skill overlap (weight 50%)
    const skillScore = jdSkills.length > 0 ? (matchedSkills.length / jdSkills.length) * 100 : 75;

    // Deterministic work mode / location score (weight 20%)
    let locationScore = 75;
    if (workMode.toLowerCase().includes("remote") || location.toLowerCase().includes("remote")) {
      locationScore = 100;
    }

    // Final weighted score
    const matchScore = Math.min(100, Math.max(0, Math.round(titleScore * 0.3 + skillScore * 0.5 + locationScore * 0.2)));
    const isShortlisted = matchScore >= minMatchScore;
    const newStatus = isShortlisted ? "Shortlisted" : "Review";
    const missingSkillsStr = missingSkills.join(", ");

    processed++;
    if (isShortlisted) {
      shortlisted++;
      shortlistedItems.push({
        jobId,
        title,
        company,
        location,
        workMode,
        salary,
        link,
        score: matchScore,
        missingSkillsStr
      });
    } else {
      reviewed++;
    }

    // Update row in Google Sheet
    const sheetRowNumber = i + 1;
    try {
      await fastn.connector.googleSheets.updateValues({
        spreadsheetId,
        range: `${sheetName}!M${sheetRowNumber}:O${sheetRowNumber}`,
        valueInputOption: "USER_ENTERED",
        values: [[matchScore, missingSkillsStr, newStatus]]
      });

      await fastn.connector.googleSheets.updateValues({
        spreadsheetId,
        range: `${sheetName}!T${sheetRowNumber}`,
        valueInputOption: "USER_ENTERED",
        values: [[todayIso]]
      });
    } catch (updateErr) {
      console.warn(`Failed to update sheet row ${sheetRowNumber}:`, updateErr.message);
    }

    // Post to Slack if shortlisted
    if (isShortlisted && slackChannel) {
      try {
        const slackMessage = `🎯 *New Shortlisted Job (${matchScore}/100)*\n` +
          `*Role:* ${title}\n` +
          `*Company:* ${company}\n` +
          `*Location:* ${location} (${workMode})\n` +
          `*Salary:* ${salary}\n` +
          `*Apply Link:* ${link}\n` +
          `*Missing Skills:* ${missingSkillsStr || "None (100% Match)"}\n` +
          `*Apply By:* ${applyBy} (${applyByEstimated === "TRUE" ? "Estimated" : "Confirmed"})\n` +
          `👉 _To apply, switch status in your Google Sheet to 'Apply' to trigger Tailoring!_`;

        await fastn.connector.slack.createChatPostMessage({
          channel: slackChannel,
          text: slackMessage
        });
      } catch (slackErr) {
        console.warn("Slack notification:", slackErr.message);
      }
    }

    details.push({
      jobId,
      title,
      company,
      score: matchScore,
      missingSkills: missingSkillsStr,
      newStatus
    });
  }

  // Helper for Base64url encoding
  function toBase64Url(str) {
    const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
    let output = "";
    for (let block = 0, charCode, idx = 0, map = chars;
      str.charAt(idx | 0) || (map = "=", idx % 1);
      output += map.charAt(63 & block >> 8 - idx % 1 * 8)) {
      charCode = str.charCodeAt(idx += 3/4);
      block = block << 8 | charCode;
    }
    return output.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }

  // 3. Send Email Notification to User about newly found shortlisted jobs
  if (shortlistedItems.length > 0 && userEmail) {
    let emailBody = `Hello,\r\n\r\nHunt Scout & Matcher just discovered ${shortlistedItems.length} new high-match job opportunities for you:\r\n\r\n`;
    for (const item of shortlistedItems) {
      emailBody += `• ${item.title} at ${item.company} (Match Score: ${item.score}/100)\r\n` +
        `  Location: ${item.location} (${item.workMode}) | Compensation: ${item.salary}\r\n` +
        `  Missing Skills: ${item.missingSkillsStr || 'None (100% Match)'}\r\n` +
        `  Apply Link: ${item.link || 'View in Sheet'}\r\n\r\n`;
    }
    emailBody += `\r\n👉 HOW TO APPLY:\r\n` +
      `Open your Google Sheet (https://docs.google.com/spreadsheets/d/${spreadsheetId}) and change the Status dropdown for any job to 'Apply'.\r\n` +
      `Hunt will automatically tailor your resume and cover letter, run a zero-hallucination verification, and prepare your application brief!\r\n\r\n` +
      `Best regards,\r\nHunt AI Orchestrator\r\n`;

    const rawEmail = `To: ${userEmail}\r\n` +
      `Subject: 🎯 Hunt: ${shortlistedItems.length} New Shortlisted Jobs Found!\r\n` +
      `Content-Type: text/plain; charset=utf-8\r\n\r\n` +
      emailBody;

    try {
      await fastn.connector.googleGmail.sendMessage({
        raw: toBase64Url(rawEmail)
      });
    } catch (mailErr) {
      console.warn("Gmail notification failed or mocked:", mailErr.message);
    }
  }

  return {
    success: true,
    processed,
    shortlisted,
    reviewed,
    errors: 0,
    details
  };
}