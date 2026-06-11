import { describe, expect, it } from "vitest";
import { buildMarketoUnsubscribeBody, findMarketoUnsubscribeForm } from "../server/marketo";

const sampleHtml = `
<script src="/js/forms2/js/forms2.min.js"></script>
<form class="mktoForm" id="mktoForm_51875">
var formDescriptor = {"Id":51875,"Vid":51875,"Status":"approved","rows":[[{"Id":957124,"Name":"Unsubscribed","Datatype":"single_checkbox"}]],"action":"\\/index.php\\/leadCapture\\/save2","munchkinId":"596-INX-704"};
`;

describe("findMarketoUnsubscribeForm", () => {
  it("extracts the form target from a preference center page", () => {
    const target = findMarketoUnsubscribeForm(sampleHtml, "https://pages.itglue.com/PreferenceCenter.html");

    expect(target).toEqual({
      submitUrl: "https://pages.itglue.com/index.php/leadCapture/save2",
      formId: "51875",
      munchkinId: "596-INX-704"
    });
  });

  it("returns null when there is no Unsubscribed field", () => {
    const html = sampleHtml.replace("Unsubscribed", "FirstName");
    expect(findMarketoUnsubscribeForm(html, "https://pages.example.com/p.html")).toBeNull();
  });

  it("returns null for non-Marketo pages", () => {
    expect(findMarketoUnsubscribeForm("<html><body>hello</body></html>", "https://example.com")).toBeNull();
  });

  it("builds the unsubscribe-all submission body", () => {
    const target = findMarketoUnsubscribeForm(sampleHtml, "https://pages.itglue.com/PreferenceCenter.html")!;
    const body = buildMarketoUnsubscribeBody("sean@example.com", target);

    expect(body).toContain("formid=51875");
    expect(body).toContain("munchkinId=596-INX-704");
    expect(body).toContain("Email=sean%40example.com");
    expect(body).toContain("Unsubscribed=yes");
  });
});
