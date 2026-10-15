export default async function (ctx) {
  const { input } = ctx;
  const action = input.action || "all";
  const spreadsheetId = input.spreadsheetId;
  if (!spreadsheetId) {
    throw new Error("Missing required input field: spreadsheetId");
  }

  const results = {};

  // 1. Run Scout (which automatically invokes Matcher & sends email alert)
  if (action === "all" || action === "scout") {
    results.scout = await fastn.flow.invoke("hunt-scout", {
      spreadsheetId,
      sheetName: input.sheetName || "Sheet1",
      targetRoles: input.targetRoles || ["Software Engineer", "Full Stack Engineer"],
      locations: input.locations || ["Remote"],
      rapidApiKey: input.rapidApiKey || "",
      resumeFileId: input.resumeFileId,
      resumeText: input.resumeText,
      minMatchScore: input.minMatchScore || 70,
      slackChannel: input.slackChannel,
      userEmail: input.userEmail || "rimahmood2020@gmail.com",
      mockJobs: input.mockJobs,
      autoMatch: true
    });
  }

  // 2. Run Matcher standalone
  if (action === "matcher") {
    results.matcher = await fastn.flow.invoke("hunt-matcher", {
      spreadsheetId,
      sheetName: input.sheetName || "Sheet1",
      resumeFileId: input.resumeFileId,
      resumeText: input.resumeText,
      minMatchScore: input.minMatchScore || 70,
      slackChannel: input.slackChannel,
      userEmail: input.userEmail || "rimahmood2020@gmail.com"
    });
  }

  // 3. Tailor resume when user toggles status to "Apply" in Sheet
  if (action === "tailor") {
    results.tailor = await fastn.flow.invoke("hunt-tailor", {
      spreadsheetId,
      sheetName: input.sheetName || "Sheet1",
      resumeFileId: input.resumeFileId,
      resumeText: input.resumeText,
      slackChannel: input.slackChannel,
      confirmApply: false
    });
  }

  // 4. Confirm Apply: Send email / portal pack & mark Applied in Sheet
  if (action === "confirm_apply") {
    results.confirmApply = await fastn.flow.invoke("hunt-tailor", {
      spreadsheetId,
      sheetName: input.sheetName || "Sheet1",
      resumeFileId: input.resumeFileId,
      resumeText: input.resumeText,
      slackChannel: input.slackChannel,
      confirmApply: true
    });
  }

  return {
    success: true,
    action,
    message: "Hunt orchestration cycle completed",
    results
  };
}