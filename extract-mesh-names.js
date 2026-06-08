#!/usr/bin/env node
/**
 * extract-mesh-names.js
 * Reads moto_web.glb (raw JSON chunk — no Draco decoding needed) and
 * regenerates parts-mesh-map.json with all mesh node names.
 *
 * Existing manual tags (bcItemNo set by workers) are preserved.
 *
 * Usage:
 *   node extract-mesh-names.js [path/to/model.glb]
 */

import { readFileSync, writeFileSync, existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const GLB_PATH  = resolve(__dirname, process.argv[2] || 'moto_web.glb');
const OUT_PATH  = resolve(__dirname, 'parts-mesh-map.json');

// GLB format: 12-byte header, then chunks.
// First chunk is always JSON (GLTF scene graph) — mesh names live here.
function extractGltfJson(path) {
  const buf       = readFileSync(path);
  const chunkLen  = buf.readUInt32LE(12);
  const chunkType = buf.readUInt32LE(16); // 0x4E4F534A = "JSON"
  if (chunkType !== 0x4E4F534A) throw new Error('GLB first chunk is not JSON');
  return JSON.parse(buf.slice(20, 20 + chunkLen).toString('utf8').replace(/\0+$/, ''));
}

function main() {
  if (!existsSync(GLB_PATH)) {
    console.error(`GLB not found: ${GLB_PATH}`);
    process.exit(1);
  }

  console.log(`Reading ${GLB_PATH}…`);
  const gltf    = extractGltfJson(GLB_PATH);
  const meshes  = gltf.meshes  || [];
  const nodes   = gltf.nodes   || [];

  // Collect node names that reference a mesh (these are what Three.js uses as mesh.name)
  const nodeNames = new Set();
  for (const node of nodes) {
    if (node.mesh !== undefined && node.name) nodeNames.add(node.name);
  }
  console.log(`Found ${nodeNames.size} mesh nodes.`);

  // Load existing map to preserve manual tags (bcItemNo set by workers)
  let existing = {};
  if (existsSync(OUT_PATH)) {
    try {
      existing = JSON.parse(readFileSync(OUT_PATH, 'utf8'));
      const tagged = Object.values(existing).filter(v => v.bcItemNo).length;
      console.log(`Loaded ${Object.keys(existing).length} existing entries (${tagged} tagged).`);
    } catch { /* fresh file if corrupt */ }
  }

  const map = {};
  let preserved = 0;

  for (const name of [...nodeNames].sort()) {
    if (existing[name]?.bcItemNo) {
      map[name] = existing[name]; // preserve worker tags
      preserved++;
    } else {
      map[name] = { bcItemNo: '', category: '' };
    }
  }

  writeFileSync(OUT_PATH, JSON.stringify(map, null, 2), 'utf8');
  console.log(`Written ${Object.keys(map).length} entries to parts-mesh-map.json`);
  console.log(`  Tagged (preserved): ${preserved}`);
  console.log(`  Untagged:           ${Object.keys(map).length - preserved}`);
}

main();
