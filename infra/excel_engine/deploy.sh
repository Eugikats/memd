#!/usr/bin/env bash
# Deploy Excel Intelligence Engine Azure resources into an existing resource group.
set -euo pipefail
RG=${1:?resource group}
LOC=${2:-eastus}
PREFIX=${3:-pulsexl}
ROOT=$(cd "$(dirname "$0")/../.." && pwd)
az deployment group create \
  --resource-group "$RG" \
  --template-file "$ROOT/infra/excel_engine/main.bicep" \
  --parameters prefix="$PREFIX" location="$LOC"
echo "Wire App Service settings from infra/excel_engine/README.md after this completes."
