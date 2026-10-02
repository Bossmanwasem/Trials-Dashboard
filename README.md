# Trials Operations Dashboard

Trials Operations Dashboard is a Windows desktop app for managing the shared Smartbox trials workflow from file intake through device preparation, QA, shipping, and lead reporting.

## Summary

The app gives the team one live workspace for tracking trial files, device assignments, shipping progress, GIPOD code usage, and daily operational reports. It is built around role-based dashboards so each user sees the tools that match their responsibilities.

## Core Features

- Shared dashboard for active trial files and lane-based prioritization.
- Device Systems Dashboard for bulk file intake, specialist claiming, file edits, GIPOD code assignment, and file action logs.
- Shipping dashboard for ready-to-ship and shipped files.
- Lead dashboard for schedules, shipped history, user training, current task summaries, weekly totals, and end-of-day cleanup reports.
- User management for admins and leads, including roles and PIN resets.
- Profile views for supported roles, including coordinator task totals and weekly schedule details.
- Automatic updates so changes made by one user appear for the rest of the team.
- Built-in update notifications for new app versions.

## Roles

- Admin: full access, user management, and role preview.
- Lead: full operational access, including lead reporting and user management.
- Device Systems Specialist: access to operational dashboards except user management.
- Device Coordinator: dashboard-focused access for claiming prep and QA work.
- Shipper: shipping dashboard access.

## Purpose

The dashboard is intended to reduce manual tracking, keep trial file status visible, and give leads a reliable view of daily throughput, shipped files, user workloads, and device training coverage.

## Self-hosted setup

The application now runs entirely on a local Node.js API and SQLite database. No cloud database account or API keys are required.

1. Install Node.js 22.5 or newer, then install dependencies with `pnpm install`.
2. Run the browser development environment with `pnpm dev`, or the Electron app with `pnpm start`.
3. Sign in with the initial administrator account: **Smartbox Admin**, PIN **7394**. Change that PIN after the first login.

The database is created automatically at `data/trials-dashboard.sqlite`. Set `TRIALS_DATA_DIR` on the server process to store it elsewhere. Back up that file (and its `-wal` file while the server is running) to preserve all dashboard data. The API listens on `127.0.0.1:47831` by default; `PORT` can override this for a standalone server and `VITE_API_URL` can point browser clients at it.
