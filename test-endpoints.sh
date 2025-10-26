#!/bin/bash

# Color codes for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo "Testing Accessibility Checker Backend Endpoints"
echo "================================================"
echo ""

BASE_URL="https://accessibility-checker-be.vercel.app"
TEST_DIR="./Accessibility Standards"

# Test files
FILES=(
  "Document Accessibility Matrix_Word.docx"
)

# Test upload endpoint
echo -e "${YELLOW}Testing Upload Endpoint...${NC}"
for file in "${FILES[@]}"; do
  echo -n "Testing: $file ... "
  RESPONSE=$(curl -s -X POST \
    -H "Content-Type: multipart/form-data" \
    -F "file=@${TEST_DIR}/${file}" \
    "${BASE_URL}/api/upload-document")
  
  if echo "$RESPONSE" | grep -q '"fileName"'; then
    echo -e "${GREEN}✓ PASSED${NC}"
  else
    echo -e "${RED}✗ FAILED${NC}"
    echo "Response: $RESPONSE"
  fi
done

echo ""
echo -e "${YELLOW}Testing Download Endpoint...${NC}"
for file in "${FILES[@]}"; do
  echo -n "Testing: $file ... "
  
  # Download the file
  OUTPUT_FILE="/tmp/remediated-$(echo "$file" | tr ' ' '_')"
  HTTP_CODE=$(curl -s -w "%{http_code}" -X POST \
    -H "Content-Type: multipart/form-data" \
    -F "file=@${TEST_DIR}/${file}" \
    "${BASE_URL}/api/download-document" \
    -o "$OUTPUT_FILE")
  
  if [ "$HTTP_CODE" = "200" ]; then
    # Verify it's a valid ZIP/DOCX file
    if unzip -t "$OUTPUT_FILE" &>/dev/null; then
      echo -e "${GREEN}✓ PASSED (Valid DOCX)${NC}"
    else
      echo -e "${RED}✗ FAILED (Invalid DOCX)${NC}"
    fi
  else
    echo -e "${RED}✗ FAILED (HTTP $HTTP_CODE)${NC}"
  fi
done

echo ""
echo -e "${YELLOW}Testing CORS Headers...${NC}"
echo -n "Testing OPTIONS preflight ... "
CORS_RESPONSE=$(curl -s -X OPTIONS \
  -H "Origin: https://accessibilitychecker25-arch.github.io" \
  -H "Access-Control-Request-Method: POST" \
  "${BASE_URL}/api/upload-document" -i)

if echo "$CORS_RESPONSE" | grep -qi "access-control-allow-origin"; then
  echo -e "${GREEN}✓ PASSED${NC}"
else
  echo -e "${RED}✗ FAILED${NC}"
fi

echo ""
echo "================================================"
echo "Test Summary Complete"
