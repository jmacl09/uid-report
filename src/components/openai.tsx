// Placeholder client for AI summary (do not put secrets here)
// If you want to use Azure OpenAI, expose a secure API route (Azure Function) and call it from here.
export type UISummaryRequest = {
  links: any[];
  associated: any[];
  kql: any;
  gdcoTickets: any[];
};

export async function summarizeUID(_input: UISummaryRequest): Promise<string> {
  // For now, the UI computes a local summary. This placeholder allows future backend wiring.
  return "";
}