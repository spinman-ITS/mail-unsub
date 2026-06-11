import { describe, expect, it } from "vitest";
import { parseUnsubscribeHeaders } from "../src/lib/headerParser";

describe("parseUnsubscribeHeaders", () => {
  it("extracts HTTPS and mailto unsubscribe methods from folded headers", () => {
    const parsed = parseUnsubscribeHeaders(
      [
        "From: Newsletter <news@example.com>",
        "List-Unsubscribe: <https://example.com/unsub?id=abc>,",
        " <mailto:unsubscribe@example.com?subject=unsubscribe>",
        "List-Unsubscribe-Post: List-Unsubscribe=One-Click"
      ].join("\r\n")
    );

    expect(parsed.oneClick).toBe(true);
    expect(parsed.httpsUrls).toEqual(["https://example.com/unsub?id=abc"]);
    expect(parsed.mailtoUrls).toEqual(["mailto:unsubscribe@example.com?subject=unsubscribe"]);
  });

  it("reports no supported method when the list unsubscribe header is missing", () => {
    const parsed = parseUnsubscribeHeaders("From: alerts@example.com\r\nSubject: Hello");

    expect(parsed.oneClick).toBe(false);
    expect(parsed.httpsUrls).toEqual([]);
    expect(parsed.mailtoUrls).toEqual([]);
  });
});
