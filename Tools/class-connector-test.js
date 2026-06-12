import { z } from "zod";

function registerConnectorTestClassTools(server) {
  // Register tool handlers (for tools/call execution)
  server.tool(
    "search_classes",
    "Search for classes matching a query. Returns matching class records with ID, name, summary, and URL.",
    {
      query: z.string().describe("Search query text"),
      top: z.number().int().min(1).max(50).optional().describe("Maximum number of results to return")
    },
    { readOnlyHint: true },
    async ({ query, top = 10 }) => ({
      content: [
        {
          type: "text",
          text: JSON.stringify({
            query,
            results: [
              {
                id: "CLASS-001",
                name: "Biology 101",
                summary: "Introductory biology class with sample test data.",
                url: "https://example.com/classes/CLASS-001"
              },
              {
                id: "CLASS-002",
                name: "Algebra I",
                summary: "Introductory algebra class with sample test data.",
                url: "https://example.com/classes/CLASS-002"
              }
            ].slice(0, top),
            total: 2,
            hasMore: false
          })
        }
      ]
    })
  );

  server.tool(
    "get_class",
    "Retrieve a single class by its unique identifier. Returns full class details.",
    {
      id: z.string().describe("The unique identifier of the class")
    },
    { readOnlyHint: true },
    async ({ id }) => ({
      content: [
        {
          type: "text",
          text: JSON.stringify({
            id,
            name: id === "CLASS-002" ? "Algebra I" : "Biology 101",
            status: "active",
            description: "Full class details from static connector test data.",
            teacher: "Sample Teacher",
            studentCount: id === "CLASS-002" ? 24 : 28,
            url: `https://example.com/classes/${encodeURIComponent(id)}`
          })
        }
      ]
    })
  );

  server.tool(
    "list_recent_classes",
    "List the most recently updated classes. Useful for checking current status and recent class activity.",
    {
      count: z.number().int().min(1).max(100).optional().describe("Number of classes to return"),
      status: z.string().optional().describe("Filter by class status: active, archived, or all")
    },
    { readOnlyHint: true },
    async ({ count = 10, status = "all" }) => ({
      content: [
        {
          type: "text",
          text: JSON.stringify({
            records: [
              {
                id: "CLASS-001",
                name: "Biology 101",
                status: "active",
                updatedAt: "2026-05-19T14:00:00Z",
                url: "https://example.com/classes/CLASS-001"
              },
              {
                id: "CLASS-002",
                name: "Algebra I",
                status: "active",
                updatedAt: "2026-05-19T13:30:00Z",
                url: "https://example.com/classes/CLASS-002"
              }
            ]
              .filter(record => status === "all" || record.status === status)
              .slice(0, count),
            total: 2,
            statusFilter: status
          })
        }
      ]
    })
  );
}

export default registerConnectorTestClassTools;
