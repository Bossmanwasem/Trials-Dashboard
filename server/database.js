import { DatabaseSync } from 'node:sqlite';
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';

const dataDirectory = process.env.TRIALS_DATA_DIR || path.resolve('data');
fs.mkdirSync(dataDirectory, { recursive: true });
export const db = new DatabaseSync(path.join(dataDirectory, 'trials-dashboard.sqlite'));
db.exec('PRAGMA journal_mode = WAL; PRAGMA foreign_keys = ON;');

db.exec(fs.readFileSync(new URL('./schema.sql', import.meta.url), 'utf8'));

const admin = db.prepare('SELECT id FROM app_users WHERE lower(first_name)=? AND lower(last_name)=?').get('smartbox', 'admin');
if (!admin) {
  db.prepare(`INSERT INTO app_users (id, first_name, last_name, role, pin_hash, created_at, updated_at)
    VALUES (?, 'Smartbox', 'Admin', 'Admin', ?, ?, ?)`).run(crypto.randomUUID(), hashPin('7394'), new Date().toISOString(), new Date().toISOString());
}

export const uuid = () => crypto.randomUUID();
export const now = () => new Date().toISOString();
export const jsonColumns = new Set(['permissions', 'weekly_schedule', 'trained_devices', 'loan_totals', 'device_totals', 'accessory_totals', 'files']);
export function decode(row) {
  if (!row) return row;
  const result = { ...row };
  for (const key of Object.keys(result)) if (jsonColumns.has(key) && typeof result[key] === 'string') result[key] = JSON.parse(result[key] || (key === 'trained_devices' || key === 'files' ? '[]' : '{}'));
  for (const key of ['is_new_hire', 'is_active', 'is_logged_in']) if (key in result) result[key] = Boolean(result[key]);
  return result;
}
export function encode(row) {
  return Object.fromEntries(Object.entries(row).map(([key, value]) => [key, jsonColumns.has(key) && typeof value !== 'string' ? JSON.stringify(value) : value]));
}

export function hashPin(pin) { const salt = crypto.randomBytes(16); return `scrypt:${salt.toString('hex')}:${crypto.scryptSync(pin, salt, 32).toString('hex')}`; }
export function verifyPin(pin, stored) { const [, salt, hash] = String(stored).split(':'); if (!salt || !hash) return false; return crypto.timingSafeEqual(Buffer.from(hash, 'hex'), crypto.scryptSync(pin, Buffer.from(salt, 'hex'), 32)); }
