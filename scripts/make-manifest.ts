import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

const domain = process.argv[2] || process.env.ADDIN_DOMAIN;

if (!domain) {
  console.error("Usage: npm run manifest:prod -- <your-app-domain>  (e.g. mail-unsub-production.up.railway.app)");
  process.exit(1);
}

const host = domain.replace(/^https?:\/\//, "").replace(/\/$/, "");
const origin = `https://${host}`;
const prodManifestId = "8c2b7f6e-41a3-4f0b-9d35-6b1f2a9c7d44";

const source = readFileSync(resolve(import.meta.dirname, "../manifest.xml"), "utf8");

const output = source
  .replaceAll("https://localhost:3003", origin)
  .replaceAll("https://localhost:8787", origin)
  .replace("<Id>1f7ad1c4-7452-4c5e-9d9d-0fe6a8a98f51</Id>", `<Id>${prodManifestId}</Id>`)
  .replace("<ProviderName>Local Dev</ProviderName>", "<ProviderName>Inman Technologies</ProviderName>")
  .replace(
    "</AppDomains>",
    "  <AppDomain>https://login.microsoftonline.com</AppDomain>\n  </AppDomains>"
  );

const target = resolve(import.meta.dirname, "../manifest.prod.xml");
writeFileSync(target, output);
console.log(`Wrote ${target} for ${origin}`);
