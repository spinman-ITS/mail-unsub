import { describe, expect, it } from "vitest";
import type { GraphMessageSummary } from "../src/lib/graph";
import { findScanCandidates } from "../src/lib/scan";

function message(overrides: Partial<GraphMessageSummary> & { id: string }): GraphMessageSummary {
  return {
    subject: "Weekly deals",
    receivedDateTime: "2026-06-10T12:00:00Z",
    from: { emailAddress: { name: "Shop", address: "deals@shop.example" } },
    internetMessageHeaders: [
      { name: "List-Unsubscribe", value: "<https://shop.example/unsub?u=1>" },
      { name: "List-Unsubscribe-Post", value: "List-Unsubscribe=One-Click" }
    ],
    ...overrides
  };
}

describe("findScanCandidates", () => {
  it("returns senders with unsubscribe headers", () => {
    const candidates = findScanCandidates([message({ id: "m1" })], []);

    expect(candidates).toHaveLength(1);
    expect(candidates[0]).toMatchObject({
      messageId: "m1",
      senderAddress: "deals@shop.example",
      senderName: "Shop"
    });
    expect(candidates[0].headers.oneClick).toBe(true);
    expect(candidates[0].headers.httpsUrls).toEqual(["https://shop.example/unsub?u=1"]);
  });

  it("skips messages without unsubscribe headers", () => {
    const plain = message({ id: "m2", internetMessageHeaders: [{ name: "Subject", value: "hi" }] });
    expect(findScanCandidates([plain], [])).toHaveLength(0);
  });

  it("skips senders already unsubscribed, case-insensitively", () => {
    const candidates = findScanCandidates([message({ id: "m3" })], ["Deals@Shop.example"]);
    expect(candidates).toHaveLength(0);
  });

  it("dedupes multiple messages from the same sender", () => {
    const candidates = findScanCandidates([message({ id: "m4" }), message({ id: "m5" })], []);
    expect(candidates).toHaveLength(1);
    expect(candidates[0].messageId).toBe("m4");
  });

  it("keeps mailto-only senders as candidates", () => {
    const mailtoOnly = message({
      id: "m6",
      internetMessageHeaders: [{ name: "List-Unsubscribe", value: "<mailto:unsub@shop.example>" }]
    });
    const candidates = findScanCandidates([mailtoOnly], []);
    expect(candidates).toHaveLength(1);
    expect(candidates[0].headers.mailtoUrls).toEqual(["mailto:unsub@shop.example"]);
  });

  it("skips messages without a sender address", () => {
    const noSender = message({ id: "m7", from: undefined });
    expect(findScanCandidates([noSender], [])).toHaveLength(0);
  });
});
