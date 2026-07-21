import { cpSync, mkdirSync, rmSync, readdirSync } from 'node:fs';
import { execSync } from 'node:child_process';
import { mkdtempSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

const repoUrl = 'https://github.com/KCLK08/buew-toolbox.git';
const workDir = mkdtempSync(join(tmpdir(), 'buew-toolbox-deploy-'));

execSync(`git clone --depth 1 --branch gh-pages ${repoUrl} ${workDir}`, {
  stdio: 'inherit'
});

for (const dir of ['tb-assistent', 'tool-platzhalter-2', 'tool-platzhalter-3', 'bautagebuch-v2']) {
  rmSync(join(workDir, dir), { recursive: true, force: true });
}

// Deploy authenticated web PWA to toolbox root (preserve tool apps).
const keep = new Set(['sitereport', 'bautagebuch', '.git']);
for (const entry of readdirSync(workDir)) {
  if (keep.has(entry)) continue;
  rmSync(join(workDir, entry), { recursive: true, force: true });
}

cpSync('web/dist', workDir, { recursive: true });
cpSync('apple-touch-icon.png', join(workDir, 'apple-touch-icon.png'));
cpSync('icon-192.png', join(workDir, 'icon-192.png'));
cpSync('icon-512.png', join(workDir, 'icon-512.png'));

rmSync(join(workDir, 'sitereport'), { recursive: true, force: true });
mkdirSync(join(workDir, 'sitereport'), { recursive: true });
cpSync('sitereport/build', join(workDir, 'sitereport'), { recursive: true });

rmSync(join(workDir, 'bautagebuch'), { recursive: true, force: true });
mkdirSync(join(workDir, 'bautagebuch'), { recursive: true });
cpSync('bautagebuch-v2/build', join(workDir, 'bautagebuch'), { recursive: true });

execSync('touch .nojekyll', { cwd: workDir, stdio: 'inherit' });
execSync('git add -A', { cwd: workDir, stdio: 'inherit' });
execSync('git commit -m "deploy: auth-enabled toolbox PWA with SiteReport and Bautagebuch"', {
  cwd: workDir,
  stdio: 'inherit'
});
execSync('git push origin gh-pages', { cwd: workDir, stdio: 'inherit' });

console.log('Deployed toolbox to gh-pages.');
