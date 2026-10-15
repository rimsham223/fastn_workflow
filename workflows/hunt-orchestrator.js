export default async function (ctx) {
  const { input } = ctx;
  const action = input.action || "all";
  const spreadsheetId = input.spreadsheetId;
  if (!spreadsheetId) {
    throw new Error("Missing required input field: spreadsheetId");
  }

  const results = {};

  // 1. Run Scout (which automatically invokes Matcher)
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
      mockJobs: input.mockJobs,
      autoMatch: true
    });
  }

  // 2. Run Matcher directly if requested separately
  if (action === "matcher") {
    results.matcher = await fastn.flow.invoke("hunt-matcher", {
      spreadsheetId,
      sheetName: input.sheetName || "Sheet1",
      resumeFileId: input.resumeFileId,
      resumeText: input.resumeText,
      minMatchScore: input.minMatchScore || 70,
      slackChannel: input.slackChannel
    });
  }

  // 3. Run Tailor on approved jobs if requested
  if (action === "tailor") {
    results.tailor = await fastn.flow.invoke("hunt-tailor", {
      spreadsheetId,
      sheetName: input.sheetName || "Sheet1",
      resumeFileId: input.resumeFileId,
      resumeText: input.resumeText,
      slackChannel: input.slackChannel,
      markApplied: Boolean(input.markApplied),
      candidateRewrite: input.candidateRewrite
    });
  }

  return {
    success: true,
    action,
    summary: "Hunt pipeline executed successfully",
    results
  };
}