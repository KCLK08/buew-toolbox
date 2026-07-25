#!/usr/bin/env node
'use strict';

/**
 * Binary-search startup crash by swapping app/_layout.tsx variants and
 * verifying whether the pre-render crash signature exists in the bundle.
 *
 * The winter-runtime crash happens BEFORE _layout executes, so every variant
 * should show the same winter signature unless metro config differs.
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const LAYOUT = path.join(ROOT, 'app/_layout.tsx');
const BACKUP = path.join(ROOT, 'app/_layout.full.tsx.bak');

const VARIANTS = {
  minimal: `import { Text, View } from 'react-native';

export default function RootLayout() {
  return (
    <View style={{ flex: 1, alignItems: 'center', justifyContent: 'center' }}>
      <Text>Minimal startup test</Text>
    </View>
  );
}
`,
  fonts: `import { View, Text, ActivityIndicator } from 'react-native';
import { useFonts, SpaceGrotesk_400Regular } from '@expo-google-fonts/space-grotesk';

export default function RootLayout() {
  const [loaded] = useFonts({ SpaceGrotesk_400Regular });
  if (!loaded) return <ActivityIndicator />;
  return <View style={{ flex: 1 }}><Text>Fonts OK</Text></View>;
}
`,
  sqlite: `import { useEffect, useState } from 'react';
import { View, Text } from 'react-native';
import { initBautagebuchDatabase } from '../src/native/bautagebuch/db/database';
import { initSiteReportDatabase } from '../src/native/sitereport/db/database';

export default function RootLayout() {
  const [ready, setReady] = useState(false);
  useEffect(() => {
    Promise.all([initBautagebuchDatabase(), initSiteReportDatabase()])
      .then(() => setReady(true))
      .catch((e) => { console.error(e); setReady(true); });
  }, []);
  return <View style={{ flex: 1 }}><Text>{ready ? 'SQLite OK' : 'Loading DB…'}</Text></View>;
}
`,
  offline: `import { View, Text } from 'react-native';
import { useOfflineBootstrap } from '../src/hooks/useOfflineBootstrap';

export default function RootLayout() {
  const { ready } = useOfflineBootstrap();
  return <View style={{ flex: 1 }}><Text>{ready ? 'Offline OK' : 'Offline…'}</Text></View>;
}
`,
  full: null
};

function analyzeBundle(bundlePath) {
  const c = fs.readFileSync(bundlePath, 'utf8');
  return {
    preRenderCrashSignature:
      /installAbortSignalPatch\)\(AbortSignal\)/.test(c) && !c.includes('resolveAbortSignalCtor(abortSignal)'),
    nullGuard: c.includes('resolveAbortSignalCtor(abortSignal)'),
    abortControllerAsset: /registerAsset\(\{[^}]*abort-controller/i.test(c),
    hasLayoutMarker: c.includes('initBautagebuchDatabase')
  };
}

function buildBundle(outFile) {
  execSync(
    `node -e "const Metro=require('metro');const c=require('./metro.config.js');Metro.runBuild(c,{entry:'node_modules/expo-router/entry.js',out:'${outFile}',platform:'android',dev:false,minify:false}).then(()=>process.exit(0)).catch(e=>{console.error(e);process.exit(1);})"`,
    { cwd: ROOT, stdio: 'pipe', maxBuffer: 50 * 1024 * 1024 }
  );
}

function restore() {
  if (fs.existsSync(BACKUP)) {
    fs.copyFileSync(BACKUP, LAYOUT);
    fs.unlinkSync(BACKUP);
  }
}

function main() {
  if (!fs.existsSync(BACKUP)) fs.copyFileSync(LAYOUT, BACKUP);

  console.log('# Binary-search startup analysis\n');
  console.log('| Variant | pre-render crash sig | null guard | abort-controller asset | layout code in bundle |');
  console.log('|---------|---------------------|------------|------------------------|----------------------|');

  for (const [name, source] of Object.entries(VARIANTS)) {
    if (name === 'full') {
      fs.copyFileSync(BACKUP, LAYOUT);
    } else {
      fs.writeFileSync(LAYOUT, source);
    }

    const out = path.join('/tmp', `layout-${name}.js`);
    try {
      buildBundle(out);
      const r = analyzeBundle(out);
      console.log(
        `| ${name} | ${r.preRenderCrashSignature ? 'YES' : 'no'} | ${r.nullGuard ? 'yes' : 'no'} | ${r.abortControllerAsset ? 'YES' : 'no'} | ${r.hasLayoutMarker ? 'yes' : 'no'} |`
      );
    } catch (err) {
      console.log(`| ${name} | BUILD FAIL | | | ${String(err.message).slice(0, 40)} |`);
    }
  }

  restore();
  console.log('\nIf pre-render crash sig is identical for minimal and full, crash is NOT in _layout initializers.');
}

main();
