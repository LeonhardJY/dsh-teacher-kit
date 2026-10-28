
import { readFileSync } from 'node:fs';

const LANGUAGE_ID = /^[a-z]{2,8}(?:-[a-z0-9]{1,8})*$/u;
const PACKAGE_NAME = /^(?:@[a-z0-9][a-z0-9._~-]*\/)?[a-z0-9][a-z0-9._~-]*$/u;
const NUMERIC = '(?:0|[1-9]\\d*)';
const PRERELEASE_PART = '(?:' + NUMERIC + '|\\d*[a-zA-Z-][0-9a-zA-Z-]*)';
const SEMVER = new RegExp('^' + NUMERIC + '\\.' + NUMERIC + '\\.' + NUMERIC + '(?:-' + PRERELEASE_PART + '(?:\\.' + PRERELEASE_PART + ')*)?(?:\\+[0-9a-zA-Z-]+(?:\\.[0-9a-zA-Z-]+)*)?$', 'u');
const COMMIT = /^[0-9a-f]{40}$/u;
const PATH_SEGMENT = /^[A-Za-z0-9._@+-]+$/u;
const DATA_ICON = /^data:(?:image\/svg\+xml|image\/png|image\/jpeg|image\/webp);base64,([A-Za-z0-9+/]*={0,2})$/u;
const BASE64 = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/u;

const errors = [];
const fail = (path, reason) => { errors.push(path + ' ' + reason); };
const isObj = (v) => typeof v === 'object' && v !== null && !Array.isArray(v);
const objectAt = (v, p) => { if (!isObj(v)) { fail(p, 'must be an object'); return {}; } return v; };
const stringAt = (v, p) => { if (typeof v !== 'string') { fail(p, 'must be a string'); return ''; } return v; };
const localized = (v, p) => {
  const text = typeof v === 'string' ? (stringAt(v, p), { en: v }) : (() => { const o = objectAt(v, p); const t = {}; for (const [k, m] of Object.entries(o)) { if (!LANGUAGE_ID.test(k)) fail(p + '.' + k, 'must be keyed by a lowercase language id'); t[k] = stringAt(m, p + '.' + k); } return t; })();
  if (!Object.hasOwn(text, 'en')) fail(p, 'must contain en');
  return text;
};
const httpUrl = (v, p, protocols) => {
  const t = stringAt(v, p);
  let u; try { u = new URL(t); } catch { fail(p, 'must be an absolute URL'); return t; }
  if (!protocols.includes(u.protocol)) fail(p, 'must use ' + protocols.join(' or '));
  if (u.username || u.password) fail(p, 'must not contain credentials');
  if (t.includes('?') || t.includes('#')) fail(p, 'must not contain a query or fragment');
  return t;
};

const indexUrl = 'https://raw.githubusercontent.com/LeonhardJY/dsh-teacher-kit/main/market/index.json';
const doc = JSON.parse(readFileSync('market/index.json', 'utf8'));
if (doc.format !== 1) fail('$.format', 'must be 1');
localized(doc.title, '$.title');
if (doc.description !== undefined) localized(doc.description, '$.description');

const md = objectAt(doc.metadata ?? {}, '$.metadata');
const declarations = (v, p, read) => {
  const out = {};
  for (const [id, member] of Object.entries(objectAt(v, p))) { if (id === '') fail(p, 'must not declare an empty id'); out[id] = read(objectAt(member, p + '.' + id), p + '.' + id); }
  return out;
};
const cats = declarations(md.categories ?? {}, '$.metadata.categories', (o, p) => (localized(o.title, p + '.title'), true));
const labels = declarations(md.labels ?? {}, '$.metadata.labels', (o, p) => (localized(o.title, p + '.title'), true));
const sortKeys = declarations(md.sortKeys ?? {}, '$.metadata.sortKeys', (o, p) => {
  if (o.order !== 'asc' && o.order !== 'desc') fail(p + '.order', 'must be "asc" or "desc"');
  localized(o.title, p + '.title'); return o.order;
});

if (!Array.isArray(doc.plugins)) fail('$.plugins', 'must be an array');
for (const [i, pl] of (doc.plugins ?? []).entries()) {
  const p = '$.plugins[' + i + ']';
  const o = objectAt(pl, p);
  const name = stringAt(o.name, p + '.name');
  if (!PACKAGE_NAME.test(name) || name.length > 214) fail(p + '.name', 'must be a valid npm package name');
  if (o.title !== undefined) localized(o.title, p + '.title');
  if (o.description !== undefined) localized(o.description, p + '.description');
  if (o.icon !== undefined) {
    const t = stringAt(o.icon, p + '.icon');
    if (t.startsWith('data:')) {
      const payload = DATA_ICON.exec(t)?.[1];
      if (!payload || !BASE64.test(payload)) fail(p + '.icon', 'must be a base64 SVG, PNG, JPEG, or WebP data URL');
      else {
        const bytes = payload.length / 4 * 3 - (payload.endsWith('==') ? 2 : payload.endsWith('=') ? 1 : 0);
        if (bytes > 32768) fail(p + '.icon', 'must not exceed 32 KiB of decoded bytes');
      }
    } else {
      const u = new URL(t, indexUrl);
      if (u.protocol !== 'https:' && (u.origin === 'null' || u.origin !== new URL(indexUrl).origin)) fail(p + '.icon', 'must resolve to https or the index origin');
    }
  }
  if (o.author !== undefined) stringAt(o.author, p + '.author');
  if (o.homepage !== undefined) httpUrl(o.homepage, p + '.homepage', ['https:', 'http:']);
  if (o.repository !== undefined) {
    const repo = objectAt(o.repository, p + '.repository');
    httpUrl(repo.url, p + '.repository.url', ['https:', 'http:']);
    if (repo.directory !== undefined) {
      const d = stringAt(repo.directory, p + '.repository.directory');
      if (d.split('/').some((s) => s === '.' || s === '..' || !PATH_SEGMENT.test(s))) fail(p + '.repository.directory', 'must be a relative directory');
    }
  }
  const pm = objectAt(o.metadata ?? {}, p + '.metadata');
  if (pm.category !== undefined && !Object.hasOwn(cats, stringAt(pm.category, p + '.metadata.category'))) fail(p + '.metadata.category', 'references undeclared id');
  for (const [j, l] of (pm.labels ?? []).entries()) if (!Object.hasOwn(labels, l)) fail(p + '.metadata.labels[' + j + ']', 'references undeclared id ' + JSON.stringify(l));
  for (const [k, v] of Object.entries(pm.sortKeys ?? {})) { if (!Object.hasOwn(sortKeys, k)) fail(p + '.metadata.sortKeys.' + k, 'references undeclared id'); if (typeof v !== 'number' || !Number.isFinite(v)) fail(p + '.metadata.sortKeys.' + k, 'must be a finite number'); }
  if (!Array.isArray(o.versions) || o.versions.length === 0) fail(p + '.versions', 'must be a non-empty array');
  const seen = new Set();
  for (const [j, ver] of (o.versions ?? []).entries()) {
    const vp = p + '.versions[' + j + ']';
    const vo = objectAt(ver, vp);
    const version = stringAt(vo.version, vp + '.version');
    if (!SEMVER.test(version)) fail(vp + '.version', 'must be an exact SemVer version');
    if (seen.has(version)) fail(vp + '.version', 'duplicates ' + version); seen.add(version);
    const ts = vo.publishTimestamp;
    if (typeof ts !== 'number' || !Number.isInteger(ts) || ts < 0 || ts >= 1e11) fail(vp + '.publishTimestamp', 'must be an integer number of Unix seconds in [0, 10^11)');
    const src = objectAt(vo.source, vp + '.source');
    if (src.type === 'npm') {
      if (src.registry !== undefined) httpUrl(src.registry, vp + '.source.registry', ['https:', 'http:']);
    } else if (src.type === 'git') {
      const commit = stringAt(src.commit, vp + '.source.commit');
      if (!COMMIT.test(commit)) fail(vp + '.source.commit', 'must be a lowercase 40-hex-digit commit SHA');
      const url = httpUrl(src.url, vp + '.source.url', ['https:']);
      if (url !== src.url) fail(vp + '.source.url', 'must be a canonical URL');
      if (!/^(?:\/[^/]+){2,}$/u.test(new URL(url).pathname)) fail(vp + '.source.url', 'must name a repository with at least two path segments');
      if (src.path !== undefined) stringAt(src.path, vp + '.source.path');
    } else fail(vp + '.source.type', 'must be "npm" or "git"');
    if (vo.engines !== undefined) {
      const e = objectAt(vo.engines, vp + '.engines');
      if (e.dsh !== undefined) {
        const range = stringAt(e.dsh, vp + '.engines.dsh');
        if (range.startsWith('workspace:')) fail(vp + '.engines.dsh', 'must not be a workspace: range');
        if (!/^[~^<>=*vxX0-9A-Za-z.|+\s-]+$/u.test(range)) fail(vp + '.engines.dsh', 'must be a SemVer range');
      }
    }
  }
}

if (errors.length) { console.log('INVALID\n' + errors.join('\n')); process.exit(1); }
console.log('VALID: format 1, ' + doc.plugins.length + ' plugin(s), icon ok, versions ok');