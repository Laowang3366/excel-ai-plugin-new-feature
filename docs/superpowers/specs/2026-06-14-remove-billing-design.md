# Remove Billing Design

## Goal

Remove paid subscription, pricing, collection settings, payment review, and entitlement gating so the add-in is free to use.

## Scope

- Delete public payment APIs and admin billing APIs.
- Remove admin routes and pages for pricing, collection settings, payment review, and subscription logs.
- Remove add-in payment/subscription pages, services, telemetry, and paid-plan copy.
- Make AI and OCR endpoints free to access for authenticated active users.
- Keep auth, user management, model settings, AI call logs, feedback, notifications, and deployment behavior.

## Data Model

Add a forward Alembic migration that drops `payment_requests`, `subscription_logs`, payment enum artifacts, and subscription/trial entitlement columns that no longer drive behavior. Keep AI call logs for observability, but remove billable semantics if they are no longer displayed or used.

## Compatibility

The API response surfaces should stop exposing paid entitlement fields to current first-party clients. If older clients call removed payment routes, they should receive normal 404 responses after the routers are removed.

## Verification

- Backend tests and lint pass.
- Admin TypeScript build/lint pass.
- Add-in TypeScript build/lint pass.
- Repository search shows no visible payment/subscription/pricing UI or active API call paths remain, except historical migrations or intentionally retained neutral usage logs.
