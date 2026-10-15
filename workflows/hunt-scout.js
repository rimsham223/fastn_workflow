export default async function (ctx) {
  const { input } = ctx;
  const spreadsheetId = input.spreadsheetId;
  if (!spreadsheetId) {
    throw new Error("Missing required input field: spreadsheetId");
  }

  const sheetName = input.sheetName || "Sheet1";
  const targetRoles = Array.isArray(input.targetRoles) && input.targetRoles.length > 0 ? input.targetRoles : ["Software Engineer"];
  const locations = Array.isArray(input.locations) && input.locations.length > 0 ? input.locations : ["Remote"];
  const rapidApiKey = input.rapidApiKey || "";

  const headers = [
    "job_id", "title", "company", "location", "work_mode", "salary", "source", "link",
    "posted_date", "days_since_posted", "apply_by", "apply_by_estimated", "match_score",
    "missing_skills", "status", "applied_on", "resume_version_link", "contact_email",
    "follow_up_due", "last_update", "notes"
  ];

  // 1. Read existing rows to dedupe by job_id
  const existingJobIds = new Set(Array.isArray(input.existingIds) ? input.existingIds : []);
  let sheetHasHeaders = false;

  try {
    const existingData = await fastn.connector.googleSheets.getValues({
      spreadsheetId,
      range: `${sheetName}!A:A`
    });
    const rows = existingData?.output?.values || existingData?.values || [];
    if (rows.length > 0) {
      sheetHasHeaders = true;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i] && rows[i][0]) {
          existingJobIds.add(String(rows[i][0]).trim());
        }
      }
    }
  } catch (err) {
    console.warn("Could not read existing Sheet rows, proceeding with provided existing IDs:", err.message);
  }

  // If sheet is empty, write headers
  if (!sheetHasHeaders && input.initHeaders !== false) {
    try {
      await fastn.connector.googleSheets.updateValues({
        spreadsheetId,
        range: `${sheetName}!A1:U1`,
        valueInputOption: "USER_ENTERED",
        values: [headers]
      });
    } catch (e) {
      console.warn("Header initialization skipped or mocked:", e.message);
    }
  }

  // 2. Fetch jobs from JSearch API or fallback to mockJobs
  const rawJobs = [];
  if (Array.isArray(input.mockJobs) && input.mockJobs.length > 0) {
    for (const j of input.mockJobs) {
      rawJobs.push(j);
    }
  } else if (rapidApiKey) {
    for (const role of targetRoles) {
      for (const loc of locations) {
        try {
          const query = `${role} in ${loc}`;
          const url = `https://jsearch.p.rapidapi.com/search?query=${encodeURIComponent(query)}&page=1&num_pages=1`;
          const res = await fetch(url, {
            method: "GET",
            headers: {
              "x-rapidapi-key": rapidApiKey,
              "x-rapidapi-host": "jsearch.p.rapidapi.com"
            }
          });
          if (res.ok) {
            const body = await res.json();
            if (body && Array.isArray(body.data)) {
              for (const item of body.data) {
                rawJobs.push(item);
              }
            }
          } else {
            console.warn(`JSearch API query '${query}' returned status:`, res.status);
          }
        } catch (fetchErr) {
          console.warn("Error calling JSearch API:", fetchErr.message);
        }
      }
    }
  }

  // 3. Normalize fields and dedupe
  let created = 0;
  let skipped = 0;
  const errorDetails = [];
  const rowsToAppend = [];
  const newJobs = [];

  const now = new Date();
  const todayIso = now.toISOString().split("T")[0];

  for (const job of rawJobs) {
    try {
      const jobId = String(job.job_id || job.id || "").trim();
      if (!jobId || existingJobIds.has(jobId)) {
        skipped++;
        continue;
      }

      // Date math
      const postedStr = job.job_posted_at_datetime_utc || job.posted_date || todayIso;
      const postedDate = new Date(postedStr);
      const validPosted = isNaN(postedDate.getTime()) ? now : postedDate;
      const daysSincePosted = Math.max(0, Math.floor((now.getTime() - validPosted.getTime()) / (1000 * 60 * 60 * 24)));

      let applyBy = "";
      let applyByEstimated = "TRUE";
      if (job.job_offer_expiration_datetime_utc) {
        applyBy = new Date(job.job_offer_expiration_datetime_utc).toISOString().split("T")[0];
        applyByEstimated = "FALSE";
      } else {
        const estDate = new Date(validPosted.getTime() + 14 * 24 * 60 * 60 * 1000);
        applyBy = estDate.toISOString().split("T")[0];
        applyByEstimated = "TRUE";
      }

      // Work mode
      let workMode = "Onsite";
      if (job.job_is_remote === true || String(job.job_title || "").toLowerCase().includes("remote")) {
        workMode = "Remote";
      } else if (job.job_employment_type) {
        workMode = job.job_employment_type;
      }

      // Salary formatting
      let salary = "Not disclosed";
      if (job.job_min_salary || job.job_max_salary) {
        const currency = job.job_salary_currency || "USD";
        const min = job.job_min_salary ? String(job.job_min_salary) : "";
        const max = job.job_max_salary ? String(job.job_max_salary) : "";
        salary = `${min}${min && max ? ' - ' : ''}${max} ${currency}`.trim();
      } else if (job.salary) {
        salary = String(job.salary);
      }

      // Location
      let location = "Remote";
      if (!job.job_is_remote) {
        const parts = [job.job_city, job.job_state, job.job_country].filter(Boolean);
        location = parts.length > 0 ? parts.join(", ") : (job.location || "Remote");
      }

      // Contact email
      const descText = String(job.job_description || job.description || "");
      const emailMatch = descText.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
      const contactEmail = emailMatch ? emailMatch[0] : (job.contact_email || "");

      const row = [
        jobId,
        String(job.job_title || job.title || "Untitled Role"),
        String(job.employer_name || job.company || "Unknown Company"),
        location,
        workMode,
        salary,
        String(job.job_publisher || job.source || "JSearch"),
        String(job.job_apply_link || job.job_google_link || job.link || ""),
        validPosted.toISOString().split("T")[0],
        daysSincePosted,
        applyBy,
        applyByEstimated,
        "", // match_score
        "", // missing_skills
        "New", // status
        "", // applied_on
        "", // resume_version_link
        contactEmail,
        "", // follow_up_due
        todayIso, // last_update
        descText.substring(0, 300) // notes
      ];

      rowsToAppend.push(row);
      existingJobIds.add(jobId);
      created++;
      newJobs.push({
        job_id: jobId,
        title: row[1],
        company: row[2],
        location: row[3],
        work_mode: row[4],
        salary: row[5],
        link: row[7],
        posted_date: row[8],
        apply_by: row[10],
        apply_by_estimated: row[11],
        status: "New"
      });
    } catch (rowErr) {
      errorDetails.push({ error: rowErr.message, item: job });
    }
  }

  // 4. Append to Google Sheets
  if (rowsToAppend.length > 0) {
    try {
      await fastn.connector.googleSheets.appendValues({
        spreadsheetId,
        range: `${sheetName}!A:U`,
        valueInputOption: "USER_ENTERED",
        insertDataOption: "INSERT_ROWS",
        values: rowsToAppend
      });
    } catch (appendErr) {
      console.warn("Append to Google Sheets failed or mocked:", appendErr.message);
    }
  }

  // 5. Automatic Chaining: Trigger Matcher automatically if new jobs were added
  let matcherResult = null;
  if (created > 0 || input.autoMatch === true) {
    try {
      matcherResult = await fastn.flow.invoke("hunt-matcher", {
        spreadsheetId,
        sheetName,
        resumeFileId: input.resumeFileId,
        resumeText: input.resumeText,
        minMatchScore: input.minMatchScore || 70,
        slackChannel: input.slackChannel
      });
    } catch (chainErr) {
      console.warn("Matcher automatic invocation:", chainErr.message);
    }
  }

  return {
    success: true,
    created,
    updated: 0,
    skipped,
    errors: errorDetails.length,
    errorDetails,
    newJobsCount: newJobs.length,
    newJobs,
    autoChainedMatcher: matcherResult
  };
}