# Accessibility Checker Tests

This directory contains comprehensive tests for the Accessibility Checker backend API, organized by functionality categories.

## 📁 Test Categories

### `/forms-and-flashing/`
Tests for form detection and flashing objects analysis:
- **`test-forms-flashing-separation.js`** - Tests separation of forms and flashing objects into distinct checks
- **`test-form-duplicate-prevention.js`** - Tests deduplication of form field detections  
- **`test-improved-form-detection.js`** - Tests enhanced form detection patterns
- **`test-forms-flashing.js`** - Legacy combined forms and flashing test

### `/location-tracking/`
Tests for location-aware accessibility detection:
- **`test-location-tracking.js`** - Tests paragraph-level location tracking system
- **`test-image-locations.js`** - Tests image accessibility with location information
- **`test-gif-location-detection.js`** - Tests GIF detection with relationship mapping and locations

### `/link-analysis/`
Tests for link accessibility analysis:
- **`test-link-descriptiveness.js`** - Tests detection of non-descriptive link text
- **`test-duplicate-links.js`** - Tests deduplication of link issues

### `/system-fixes/`
Tests for system-level fixes and improvements:
- **`test-flagging-system.js`** - Tests conversion from auto-fix to flagging system
- **`test-function-fix.js`** - Tests fix for function definition scope issues

### `/legacy/`
Historical tests from earlier development phases:
- **`test-advanced-shadows.js`** - Advanced shadow removal tests
- **`test-batch-processing.js`** - Batch document processing tests
- **`test-comprehensive-shadows.js`** - Comprehensive shadow detection tests
- **`test-line-spacing-detection.js`** - Line spacing analysis tests
- **`test-line-spacing.js`** - Line spacing validation tests
- **`test-session-system.js`** - Session management tests
- **`test-shadow-removal.js`** - Shadow removal functionality tests
- **`test-simple-detection.js`** - Basic detection tests

## 🚀 Running Tests

### Run All Tests
```bash
# Run all tests in a category
node tests/forms-and-flashing/test-forms-flashing-separation.js
node tests/location-tracking/test-gif-location-detection.js
node tests/link-analysis/test-duplicate-links.js
node tests/system-fixes/test-function-fix.js
```

### Run Category Tests
```bash
# Forms and Flashing
find tests/forms-and-flashing -name "*.js" -exec node {} \;

# Location Tracking  
find tests/location-tracking -name "*.js" -exec node {} \;

# Link Analysis
find tests/link-analysis -name "*.js" -exec node {} \;

# System Fixes
find tests/system-fixes -name "*.js" -exec node {} \;
```

## 📊 Test Coverage

### Current Functionality Coverage:
- ✅ **Forms Detection** - Enhanced patterns, duplicate prevention, location tracking
- ✅ **Flashing Objects** - Separate analysis, animation detection, location tracking  
- ✅ **GIF Detection** - Location mapping, relationship parsing, accessibility recommendations
- ✅ **Link Analysis** - Non-descriptive detection, pattern matching, deduplication
- ✅ **Location Tracking** - Paragraph-level precision, page estimation, context tracking
- ✅ **System Integrity** - Function scoping, error handling, flagging system

### Key Features Tested:
- **Separation of Concerns** - Forms and flashing objects as distinct checks
- **Duplicate Prevention** - Location-based deduplication using Set()
- **Priority Selection** - Smart form type prioritization when multiple patterns match
- **Comprehensive Detection** - 14+ form element types, 9+ animation types
- **Location Precision** - Paragraph numbers, page estimates, heading context
- **Error Handling** - Function definition fixes, scope resolution

## 🎯 Development Workflow

These tests were created during active development to validate:

1. **API Response Structure** - Ensuring proper JSON formatting and field consistency
2. **Location Information** - Verifying paragraph-level tracking across all checks
3. **Deduplication Logic** - Preventing multiple reports of the same accessibility issue
4. **Pattern Matching** - Validating regex patterns for various document structures
5. **Function Integration** - Testing function availability and proper scoping
6. **Feature Separation** - Ensuring modular, maintainable code architecture

## 📈 Success Metrics

Tests validate these success criteria:
- **No Duplicate Reports** - Each accessibility issue reported exactly once
- **Accurate Locations** - Precise paragraph and page information for all issues  
- **Comprehensive Coverage** - Detection of all supported accessibility issues
- **Error-Free Execution** - No function definition or runtime errors
- **Consistent Formatting** - Uniform API response structure across all checks

---

**Last Updated**: November 12, 2025  
**Total Tests**: 20+ comprehensive test files  
**Coverage**: Forms, Flashing Objects, GIFs, Links, Locations, System Fixes