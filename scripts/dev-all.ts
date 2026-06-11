import { spawn } from "node:child_process";

const children = [
  spawn("npm", ["run", "dev"], { stdio: "inherit" }),
  spawn("npm", ["run", "server"], { stdio: "inherit", env: { ...process.env, NODE_ENV: "development" } })
];

function stopAll() {
  for (const child of children) {
    if (!child.killed) {
      child.kill("SIGTERM");
    }
  }
}

for (const child of children) {
  child.on("exit", (code) => {
    if (code && code !== 0) {
      stopAll();
      process.exitCode = code;
    }
  });
}

process.on("SIGINT", () => {
  stopAll();
  process.exit(130);
});

process.on("SIGTERM", () => {
  stopAll();
  process.exit(143);
});
