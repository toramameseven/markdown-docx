const esbuild = require("esbuild");
const fs = require("fs");

//esbuild ./src/extension.ts --bundle --outfile=out/extension.js --external:vscode --format=cjs --platform=node"

const jsdomPatch = {
    name: 'jsdom-patch',
    setup(build) {
        build.onLoad({ filter: /xmlhttprequest\.js$/ }, async (args) => {
            let contents = await fs.promises.readFile(args.path, 'utf8');
            contents = contents.replace(
                'const syncWorkerFile = require.resolve ? require.resolve("./xhr-sync-worker.js") : null;',
                `const syncWorkerFile = "${require.resolve('jsdom/lib/jsdom/living/xhr/xhr-sync-worker.js')}";`.replaceAll('\\', process.platform === 'win32' ? '\\\\' : '\\'),
            );
            return { contents, loader: 'js' };
        });
    },
};;

esbuild.build({
  entryPoints: ["./src/extension.ts"],
  bundle: true,
  outfile: "out/extension.js",
  external: ["vscode"],
  format: "cjs",
  platform: "node",
  plugins: [jsdomPatch],
  sourcemap: true,
});
