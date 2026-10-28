import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

// Nombre del repositorio: en GitHub Actions viene de GITHUB_REPOSITORY ("usuario/repo").
// Así el mismo código despliega bien en producción y en un repo de pruebas.
const PROD_REPO = "mi-proyecto-tailwindv2";
const REPO = process.env.GITHUB_REPOSITORY?.split("/")[1] || PROD_REPO;
// Cualquier repo distinto del de producción se marca como entorno de pruebas.
const ENV_LABEL = REPO === PROD_REPO ? "" : "PRUEBAS";

// Identificador único de cada build. En GitHub Actions usa el SHA del commit;
// en local, la fecha/hora del build.
const BUILD_ID = process.env.GITHUB_SHA?.slice(0, 12) || String(Date.now());

// Publica dist/version.json con el mismo BUILD_ID que queda embebido en el bundle.
// La app lo consulta para avisar cuando hay una versión nueva desplegada.
const versionFile = (): Plugin => ({
  name: "version-file",
  apply: "build",
  generateBundle() {
    this.emitFile({
      type: "asset",
      fileName: "version.json",
      source: JSON.stringify({ version: BUILD_ID }),
    });
  },
});

export default defineConfig({
  plugins: [react(), tailwindcss(), versionFile()],
  base: `/${REPO}/`,
  define: {
    __BUILD_ID__: JSON.stringify(BUILD_ID),
    __ENV_LABEL__: JSON.stringify(ENV_LABEL),
  },
});
