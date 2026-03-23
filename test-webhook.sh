#!/bin/bash
# Test the webhook endpoint with a sample address.
# Usage: ./test-webhook.sh [address]

ADDRESS="${1:-Testveien 1, 0001 Oslo}"
PORT="${WEBHOOK_PORT:-3000}"
SECRET="${WEBHOOK_SECRET:-}"

AUTH_HEADER=""
if [ -n "$SECRET" ]; then
  AUTH_HEADER="-H \"Authorization: Bearer $SECRET\""
fi

echo "Sending webhook request to http://localhost:$PORT/webhook"
echo "Address: $ADDRESS"
echo "---"

curl -s -X POST "http://localhost:$PORT/webhook" \
  -H "Content-Type: application/json" \
  ${AUTH_HEADER:+-H "Authorization: Bearer $SECRET"} \
  -d "{\"address\": \"$ADDRESS\"}" | python -m json.tool 2>/dev/null || echo "(raw output above)"
