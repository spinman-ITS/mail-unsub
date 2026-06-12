import { describe, expect, it } from "vitest";
import { isPreferenceCenter } from "../server/marketo";

const sampleHtml = `
<script src="/js/forms2/js/forms2.min.js"></script>
<form class="mktoForm" id="mktoForm_51875">
var formDescriptor = {
  "Id": 51875,
  "Vid": 51875,
  "rows": [
    [{"Id":957121,"Name":"Email","Datatype":"email"}],
    [{"Id":957124,"Name":"Unsubscribed","Datatype":"single_checkbox","PicklistValues":[{"label":"","value":"yes"}]}]
  ],
  "action": "\\/index.php\\/leadCapture\\/save2",
  "munchkinId": "596-INX-704"
};
`;

describe("isPreferenceCenter", () => {
  it("detects Marketo preference centers", () => {
    expect(isPreferenceCenter(sampleHtml)).toBe(true);
  });

  it("returns false when the form descriptor is missing", () => {
    expect(isPreferenceCenter('<html><body>"Name":"Unsubscribed"</body></html>')).toBe(false);
  });

  it("returns false without the Unsubscribed field", () => {
    expect(isPreferenceCenter(sampleHtml.replace(/"Unsubscribed"/g, '"FirstName"'))).toBe(false);
  });

  it("returns false for non-Marketo pages", () => {
    expect(isPreferenceCenter("<html><body>hello</body></html>")).toBe(false);
  });
});
