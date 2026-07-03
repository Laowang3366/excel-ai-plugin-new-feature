# Remove Billing Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove the paid billing/subscription system and make the plugin free.

**Architecture:** Dedicated payment/subscription modules are deleted. Shared user/account and usage code is simplified so AI/OCR requests stay authenticated but no longer check entitlement or charge billable time. Frontend navigation and contracts remove billing pages and paid copy.

**Tech Stack:** FastAPI, SQLAlchemy/Alembic, React, TypeScript, Vite.

---

### Task 1: Backend Billing Removal

**Files:**
- Modify: `backend/app/main.py`
- Modify: `backend/app/api/admin.py`
- Modify: `backend/app/services/usage_service.py`
- Modify: `backend/app/api/ocr.py`
- Modify: `backend/app/api/ocr_upload_policy.py`
- Modify: `backend/app/schemas/auth.py`
- Modify: `backend/app/services/user_info_projection.py`
- Modify: `backend/app/schemas/admin.py`
- Modify: `backend/app/api/admin_user_projections.py`
- Delete: `backend/app/api/payment.py`
- Delete: `backend/app/api/payment_order_projections.py`
- Delete: `backend/app/api/admin_payment_review_events.py`
- Delete: `backend/app/api/admin_subscription_log_projections.py`
- Delete: `backend/app/models/payment_request.py`
- Delete: `backend/app/models/subscription_log.py`
- Delete: `backend/app/schemas/payment.py`
- Delete: `backend/app/services/payment_service.py`
- Delete: `backend/app/services/payment_settings_service.py`
- Delete: `backend/app/services/payment_order_transition_policy.py`
- Delete: `backend/app/services/payment_public_order_action_policy.py`
- Delete: `backend/app/services/pricing_service.py`
- Delete: `backend/app/services/subscription_service.py`

- [x] Add backend tests that assert health/import still works and removed payment endpoints are absent.
- [x] Remove payment router and admin billing route blocks.
- [x] Replace quota/subscription checks with free authenticated access.
- [x] Keep AI call logging but stop charging billable time.
- [x] Run backend pytest and ruff.

### Task 2: Database Migration

**Files:**
- Modify: `backend/alembic/env.py`
- Create: `backend/alembic/versions/0014_remove_billing.py`

- [x] Remove deleted billing model imports from Alembic metadata setup.
- [x] Add a forward migration that drops payment/subscription tables and billing columns.
- [x] Ensure downgrade restores schema enough for local rollback.

### Task 3: Admin UI Cleanup

**Files:**
- Modify: `admin/src/components/adminShellMenu.tsx`
- Modify: `admin/src/components/adminShellRoutes.tsx`
- Modify: `admin/src/services/adminContracts.ts`
- Modify: `admin/src/services/adminService.ts`
- Modify: `admin/src/pages/UserDetailPage.tsx`
- Modify: `admin/src/pages/users/*`
- Modify: `admin/src/pages/dashboard/dashboardStatsProjection.tsx`
- Delete: `admin/src/pages/PricingPage.tsx`
- Delete: `admin/src/pages/pricing/*`
- Delete: `admin/src/pages/PaymentSettingsPage.tsx`
- Delete: `admin/src/pages/PaymentReviewPage.tsx`
- Delete: `admin/src/pages/payment/*`
- Delete: `admin/src/pages/SubscriptionLogsPage.tsx`
- Delete: `admin/src/pages/subscriptions/*`
- Delete: `admin/src/components/SubscriptionModal.tsx`

- [x] Remove billing menu entries and route imports.
- [x] Remove user subscription actions and subscription log panels.
- [x] Remove billing contracts and service methods.
- [x] Run admin TypeScript build/lint.

### Task 4: Add-in Cleanup

**Files:**
- Modify: `add-in/src/pages/LoginPage.tsx`
- Modify: `add-in/src/pages/OcrPage.tsx`
- Modify: `add-in/src/features/ocr/ocrApiErrorMessage.ts`
- Modify: `add-in/src/pages/SettingsPage.tsx`
- Modify: `add-in/src/services/authService.ts`
- Delete: `add-in/src/services/subscriptionService.ts`
- Delete: `add-in/src/pages/PaymentPage.tsx`
- Delete: `add-in/src/pages/SubscriptionPage.tsx`
- Delete: `add-in/src/features/payment/*`
- Delete: `add-in/src/features/subscription/*`
- Delete: `add-in/src/features/settings/settingsSubscriptionOverviewProjection.ts`

- [x] Remove orphaned payment/subscription source.
- [x] Replace paid/trial copy with free-account copy.
- [x] Run add-in TypeScript build/lint.

### Task 5: Final Verification

- [x] Search for active payment/subscription/pricing references.
- [x] Run backend tests.
- [x] Run admin lint/build.
- [x] Run add-in lint/build.
- [ ] Commit the complete cleanup on `codex/remove-billing-and-subscriptions`.
