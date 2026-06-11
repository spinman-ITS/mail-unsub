import { afterEach, describe, expect, it, vi } from "vitest";
import { hasOutlookMessageContext, waitForOutlookMessageContext } from "../src/office/host";

describe("hasOutlookMessageContext", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("returns false when the add-in page is opened outside Outlook", () => {
    vi.stubGlobal("Office", undefined);

    expect(hasOutlookMessageContext()).toBe(false);
  });

  it("returns true when Outlook exposes message header APIs", () => {
    vi.stubGlobal("Office", {
      context: {
        mailbox: {
          item: {
            getAllInternetHeadersAsync: vi.fn()
          }
        }
      }
    });

    expect(hasOutlookMessageContext()).toBe(true);
  });

  it("waits for Office to become ready before checking message APIs", async () => {
    vi.stubGlobal("Office", {
      onReady: vi.fn(async () => undefined),
      context: {
        mailbox: {
          item: {
            getAllInternetHeadersAsync: vi.fn()
          }
        }
      }
    });

    await expect(waitForOutlookMessageContext()).resolves.toBe(true);
  });
});
