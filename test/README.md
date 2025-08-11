# OCRFunctions Testing Guide

This document explains how to run and use the Pester tests for the OCRFunctions.ps1 module.

## Prerequisites

- PowerShell 5.1 or later
- Pester module (will be automatically installed if not present)
- Windows 10/11 with OCR capabilities

## Test Files

- `OCRFunctions.Tests.ps1` - Main test file containing all test cases
- `Run-Tests.ps1` - Test runner script with various options
- `TestData/` - Directory for test images and temporary files

## Running Tests

### Basic Test Execution

```powershell
# Navigate to tests directory
cd test

# Run all tests with default settings
.\Run-Tests.ps1

# Run tests with detailed output
.\Run-Tests.ps1 -PassThru

# Run specific test categories
.\Run-Tests.ps1 -Tag "UtilityFunctions"
```

### Advanced Test Options

```powershell
# Run with code coverage analysis
.\Run-Tests.ps1 -CodeCoverage

# Export results to specific format
.\Run-Tests.ps1 -OutputFormat "NUnitXml" -OutputFile "MyTestResults.xml"

# Skip certain test categories
.\Run-Tests.ps1 -ExcludeTag "Integration"
```

### Direct Pester Execution

```powershell
# Install Pester if needed
Install-Module -Name Pester -Force -SkipPublisherCheck

# Run tests directly from tests directory
Invoke-Pester -Path ".\OCRFunctions.Tests.ps1"
```

### Running from Root Directory

```powershell
# From project root, you can also run:
Invoke-Pester -Path ".\tests\OCRFunctions.Tests.ps1"

# Or use the test runner from root:
.\tests\Run-Tests.ps1
```

## Test Categories

### 1. Utility Functions
Tests for basic helper functions:
- `Get-CurrentProcessId` - Process ID retrieval
- `Get-ParentProcessId` - Parent process identification
- `Find-MSEdgeWebView2Process` - WebView process discovery

### 2. Window Management
Tests for window interaction functions:
- `Find-WindowByTitle` - Window enumeration and search
- `Activate-Window` - Window activation
- `Show-Process` - Process window management

### 3. Input and Screenshot Functions
Tests for system interaction:
- `Click-Coordinates` - Mouse click simulation
- `Take-Screenshot` - Screen capture functionality
- `Send-Text` - Keyboard input simulation

### 4. OCR Functions
Tests for optical character recognition:
- `Find-TextInImageUsingWindowsOCR` - Text search in images
- `Get-AllTextFromImage` - Full text extraction

### 5. Integration Tests
End-to-end workflow testing and error handling scenarios.

## Directory Structure

```
uwc_custom/
├── OCRFunctions.ps1           # Main module
├── tests/                     # Test directory
│   ├── OCRFunctions.Tests.ps1 # Test cases
│   ├── Run-Tests.ps1          # Test runner
│   ├── README.md              # This file
│   └── TestData/              # Auto-generated test data
│       ├── test-image.png     # Minimal test image
│       └── (temp files)       # Temporary test files
```

## Test Data

### Automatically Generated
The test suite creates minimal test data:
- `TestData/test-image.png` - Minimal 1x1 pixel PNG for basic OCR testing
- Temporary screenshot files for integration testing

### Custom Test Images
To test with real OCR scenarios, place image files in the `TestData/` directory:
- Use PNG, JPEG, or other supported formats
- Include images with various text content
- Test with different languages and fonts

## Understanding Test Results

### Test States
- **Passed** ✅ - Test executed successfully
- **Failed** ❌ - Test did not meet expectations
- **Skipped** ⏭️ - Test was bypassed (usually due to missing prerequisites)

### Common Skip Reasons
- No main window handle available for process tests
- Test images not accessible
- OCR engine not available for current language

### Expected Behaviors
Some tests are designed to handle graceful failures:
- Invalid window handles should return `$false`
- Missing image files should throw appropriate exceptions
- Non-existent processes should be handled gracefully

## Troubleshooting

### OCR Tests Failing
1. Ensure Windows OCR is enabled:
   - Windows Settings > Apps > Optional features
   - Add "Optical character recognition" if missing

2. Check language support:
   - OCR engine needs at least one supported language
   - English is typically included by default

### Permission Issues
Some tests require elevated privileges or user interaction:
- Screenshot functions need display access
- Input simulation may require active desktop session
- Window management needs appropriate permissions

### Mock Limitations
Current test implementation has limited mocking for:
- Windows Runtime OCR APIs
- Win32 API functions
- Hardware-dependent operations

## Extending Tests

### Adding New Test Cases

```powershell
Describe "My New Feature" {
    Context "Specific Scenario" {
        It "Should behave as expected" {
            # Arrange
            $input = "test data"
            
            # Act
            $result = My-Function -Parameter $input
            
            # Assert
            $result | Should -Be "expected output"
        }
    }
}
```

### Best Practices
- Use descriptive test names
- Include both positive and negative test cases
- Test parameter validation
- Handle expected exceptions appropriately
- Clean up temporary files in `AfterEach`/`AfterAll`

### Performance Considerations
- OCR tests may be slow on some systems
- Screenshot operations require active display
- Large test images impact execution time

## Continuous Integration

### GitHub Actions Example
```yaml
- name: Run Pester Tests
  run: |
    cd tests
    .\Run-Tests.ps1 -OutputFormat "NUnitXml" -OutputFile "TestResults.xml"
  shell: powershell

- name: Publish Test Results
  uses: dorny/test-reporter@v1
  if: always()
  with:
    name: PowerShell Tests
    path: tests/TestResults.xml
    reporter: dotnet-trx
```

## Contributing

When adding new functions to OCRFunctions.ps1:
1. Add corresponding test cases to `OCRFunctions.Tests.ps1`
2. Include parameter validation tests
3. Test error conditions
4. Update this documentation
5. Ensure all tests pass before submitting

## Support

For test-related issues:
1. Check test output for specific failure details
2. Verify prerequisites are met
3. Review function documentation in the main module
4. Test individual functions in isolation