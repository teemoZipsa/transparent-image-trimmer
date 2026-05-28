import { readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { defineConfig } from "vite";

const certsDirectory = join(homedir(), ".office-addin-dev-certs");

export default defineConfig({
  server: {
    host: "localhost",
    port: 3000,
    https: {
      ca: readFileSync(join(certsDirectory, "ca.crt")),
      cert: readFileSync(join(certsDirectory, "localhost.crt")),
      key: readFileSync(join(certsDirectory, "localhost.key"))
    }
  }
});
