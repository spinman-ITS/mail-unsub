import { describe, expect, it } from "vitest";
import { findBodyUnsubscribeLinks } from "../src/lib/bodyLinks";

describe("findBodyUnsubscribeLinks", () => {
  it("finds visible unsubscribe links and ranks them above preference pages", () => {
    const links = findBodyUnsubscribeLinks(`
      <p>This message was sent to you.</p>
      <a href="https://example.com/preferences">Privacy</a>
      <a href="https://example.com/unsubscribe?id=abc">Unsubscribe</a>
    `);

    expect(links).toEqual([
      { url: "https://example.com/unsubscribe?id=abc", label: "Unsubscribe" },
      { url: "https://example.com/preferences", label: "Privacy" }
    ]);
  });

  it("finds preference center links like Marketo subscription pages", () => {
    const links = findBodyUnsubscribeLinks(`
      <a href="https://pages.itglue.com/PreferenceCenter.html">Update your preferences</a>
    `);

    expect(links).toEqual([
      { url: "https://pages.itglue.com/PreferenceCenter.html", label: "Update your preferences" }
    ]);
  });

  it("finds rewritten unsubscribe URLs when the link text says click here", () => {
    const links = findBodyUnsubscribeLinks(`
      To unsubscribe <a href="https://protect.checkpoint.com/v2/r01/___https://sender.example.com/communication/subscribe?id=123___.abc">click here</a>.
    `);

    expect(links).toEqual([
      {
        url: "https://protect.checkpoint.com/v2/r01/___https://sender.example.com/communication/subscribe?id=123___.abc",
        label: "click here"
      }
    ]);
  });

  it("finds security-rewritten opaque links when the sentence is outside the anchor's parent", () => {
    const links = findBodyUnsubscribeLinks(`
      <td><p>To unsubscribe from emails, <span><a href="https://protect.example-scanner.com/s/AbC123xyz">click here</a></span>.</p></td>
    `);

    expect(links).toEqual([
      { url: "https://protect.example-scanner.com/s/AbC123xyz", label: "click here" }
    ]);
  });

  it("ignores unrelated links", () => {
    const links = findBodyUnsubscribeLinks('<a href="https://example.com/register">Register</a>');

    expect(links).toEqual([]);
  });
});
