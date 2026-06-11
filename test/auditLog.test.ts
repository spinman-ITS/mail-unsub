import { describe, expect, it } from "vitest";
import { redactUrl } from "../server/auditLog";

describe("redactUrl", () => {
  it("keeps useful routing details while redacting query tokens", () => {
    expect(redactUrl("https://example.com/unsubscribe/path?id=secret&token=abc")).toBe(
      "https://example.com/unsubscribe/path?..."
    );
  });

  it("handles invalid URLs safely", () => {
    expect(redactUrl("not a url")).toBe("invalid-url");
  });
});
