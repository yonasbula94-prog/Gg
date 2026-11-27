#!/usr/bin/env bash
# Deploys the app to Cloud Run with a Cloud SQL (PostgreSQL) database.
# Run from the repo root after `gcloud auth login`.
#
# Required env vars (set before running):
#   SQL_ADMIN_PASSWORD   password for the new Postgres user
# Optional:
#   VENDOR_ADMIN_EMAILS  comma-separated admin emails
#   REGION               default us-central1

set -euo pipefail

PROJECT_ID="gen-lang-client-0284106792"
REGION="${REGION:-us-central1}"
SERVICE="grand-group-procurement"
SQL_INSTANCE="procurement-sql"
SQL_DB="procurement"
SQL_USER="procurement_app"

: "${SQL_ADMIN_PASSWORD:?Set SQL_ADMIN_PASSWORD first}"

gcloud config set project "$PROJECT_ID"
gcloud services enable run.googleapis.com sqladmin.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com

# 1. Cloud SQL instance (skipped if it already exists)
if ! gcloud sql instances describe "$SQL_INSTANCE" >/dev/null 2>&1; then
  gcloud sql instances create "$SQL_INSTANCE" \
    --database-version=POSTGRES_16 \
    --tier=db-f1-micro \
    --region="$REGION" \
    --storage-size=10
fi

# 2. Database and user
gcloud sql databases create "$SQL_DB" --instance="$SQL_INSTANCE" 2>/dev/null || true
gcloud sql users create "$SQL_USER" --instance="$SQL_INSTANCE" --password="$SQL_ADMIN_PASSWORD" 2>/dev/null \
  || gcloud sql users set-password "$SQL_USER" --instance="$SQL_INSTANCE" --password="$SQL_ADMIN_PASSWORD"

CONNECTION_NAME="$(gcloud sql instances describe "$SQL_INSTANCE" --format='value(connectionName)')"

# 3. Build with Cloud Build and deploy to Cloud Run.
#    SQL_HOST is the Cloud SQL unix socket path that Cloud Run mounts.
ENV_VARS="SQL_HOST=/cloudsql/${CONNECTION_NAME},SQL_DB_NAME=${SQL_DB},SQL_ADMIN_USER=${SQL_USER},SQL_ADMIN_PASSWORD=${SQL_ADMIN_PASSWORD}"
if [ -n "${VENDOR_ADMIN_EMAILS:-}" ]; then
  ENV_VARS="${ENV_VARS},VENDOR_ADMIN_EMAILS=${VENDOR_ADMIN_EMAILS}"
fi

gcloud run deploy "$SERVICE" \
  --source . \
  --region "$REGION" \
  --allow-unauthenticated \
  --add-cloudsql-instances "$CONNECTION_NAME" \
  --set-env-vars "$ENV_VARS" \
  --memory 1Gi

gcloud run services describe "$SERVICE" --region "$REGION" --format='value(status.url)'
echo "Add the URL's host to Firebase Auth > Settings > Authorized domains."
