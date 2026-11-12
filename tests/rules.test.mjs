// اختبار قواعد قاعدة البيانات في محاكي Firebase: الصلاحيات، وأن سجل التدقيق لا يُعدّل ولا يُحذف حتى من المدير.
// التشغيل: npx firebase emulators:exec --only database --project demo-jb "node tests/rules.test.mjs database.rules.json"
import { initializeTestEnvironment, assertSucceeds, assertFails } from '@firebase/rules-unit-testing';
import fs from 'fs';
const rules = fs.readFileSync(process.argv[2], 'utf8');
const env = await initializeTestEnvironment({ projectId: 'demo-jb', database: { rules, host: '127.0.0.1', port: 9000 } });
await env.withSecurityRulesDisabled(async c => {
  await c.database().ref('jabeenBookings').set({
    users: { 'admin@x,com': { role: 'admin', name: 'A' }, 'maker@x,com': { role: 'maker', name: 'M' } },
    audit: { a1: { at: 1, whoEmail: 'maker@x.com', action: 'old' } },
    master: { U1: { price: 1 } }, config: { mail: { on: false } } });
});
const as = (email) => env.authenticatedContext(email, { email, email_verified: true }).database().ref('jabeenBookings');
const admin = as('admin@x.com'), maker = as('maker@x.com'), stranger = as('evil@x.com');
const r = [];
const t = async (name, p, ok) => { try { await (ok ? assertSucceeds(p) : assertFails(p)); r.push('PASS ' + name); } catch (e) { r.push('FAIL ' + name + ' :: ' + e.message.slice(0, 120)); } };
await t('admin writes config', admin.child('config/mail/on').set(true), true);
await t('admin writes master', admin.child('master/U1/sakaniCode').set('S1'), true);
await t('admin multi-path update', admin.update({ 'master/U1/x': 1, 'sakaniSync/last': 2, 'idx/unit/U1': 'b1' }), true);
await t('admin writes users', admin.child('users/new@x,com').set({ role: 'maker' }), true);
await t('admin writes bookings', admin.child('bookings/b9').set({ status: 'cancelled' }), true);
await t('admin writes customers', admin.child('customers/1012345678/notes/n1').set({ t: 1 }), true);
await t('admin writes mailLog/pledges/backupStatus', admin.update({ 'mailLog/b1/x': 1, 'pledges/b1': { a: 1 }, 'backupStatus/x': 1 }), true);
await t('admin appends audit', admin.child('audit').push({ at: 2, whoEmail: 'admin@x.com', action: 'x' }), true);
await t('admin CANNOT edit audit', admin.child('audit/a1/action').set('tampered'), false);
await t('admin CANNOT delete audit entry', admin.child('audit/a1').remove(), false);
await t('admin CANNOT wipe audit', admin.child('audit').remove(), false);
await t('admin CANNOT replace root', admin.set({}), false);
await t('maker appends audit', maker.child('audit').push({ at: 3, whoEmail: 'maker@x.com', action: 'y' }), true);
await t('maker CANNOT write config', maker.child('config/mail/on').set(false), false);
await t('stranger CANNOT read', stranger.child('bookings').once('value'), false);
await t('admin reads', admin.child('audit').once('value'), true);
console.log(r.join('\n'));
await env.cleanup(); process.exit(r.some(x => x.startsWith('FAIL')) ? 1 : 0);
