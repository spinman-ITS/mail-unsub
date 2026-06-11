import { describe, expect, it } from "vitest";
import { isSafeRedirectTarget, validateUnsubscribeUrl } from "../server/safety";

describe("validateUnsubscribeUrl", () => {
  it("allows public HTTPS unsubscribe URLs", async () => {
    await expect(validateUnsubscribeUrl("https://example.com/unsubscribe?id=123")).resolves.toEqual({
      ok: true,
      url: "https://example.com/unsubscribe?id=123"
    });
  });

  it.each([
    "http://example.com/unsubscribe",
    "https://localhost/unsubscribe",
    "https://127.0.0.1/unsubscribe",
    "https://10.0.0.1/unsubscribe",
    "https://192.168.1.10/unsubscribe",
    "ftp://example.com/unsubscribe"
  ])("rejects unsafe URL %s", async (url) => {
    await expect(validateUnsubscribeUrl(url)).resolves.toMatchObject({ ok: false });
  });
});

describe("isSafeRedirectTarget", () => {
  it("allows public HTTP redirects from legacy unsubscribe systems", () => {
    expect(isSafeRedirectTarget("http://eastparkerchamber.chambermaster.com/communication/subscribe?id=123")).toBe(true);
  });

  it.each(["http://localhost/unsubscribe", "http://127.0.0.1/unsubscribe", "http://10.0.0.1/unsubscribe", "ftp://example.com/unsubscribe"])(
    "rejects unsafe redirect %s",
    (url) => {
      expect(isSafeRedirectTarget(url)).toBe(false);
    }
  );
});
