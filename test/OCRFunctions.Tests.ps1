# OCRFunctions.Tests.ps1
# Pester tests for OCRFunctions.ps1

BeforeAll {
    # Import the module being tested (now one level up from tests directory)
    . "$PSScriptRoot\..\OCRFunctions.ps1"
    
    # Create test directories and files if they don't exist
    $TestDataPath = Join-Path $PSScriptRoot "TestData"
    if (-not (Test-Path $TestDataPath)) {
        New-Item -ItemType Directory -Path $TestDataPath -Force
    }
    
    # Create a simple test image for OCR testing (1x1 white pixel PNG)
    $TestImagePath = Join-Path $TestDataPath "test-image.png"
    if (-not (Test-Path $TestImagePath)) {
        # Create minimal PNG data for testing
        $pngData = @(
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,  # PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,  # IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,  # 1x1 dimensions
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,  # bit depth, color type, etc.
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,  # IDAT chunk
            0x54, 0x08, 0x99, 0x01, 0x01, 0x03, 0x00, 0xFC,
            0xFF, 0xFF, 0xFF, 0xFF, 0x02, 0x00, 0x01, 0xE2,
            0x21, 0xBC, 0x33, 0x00, 0x00, 0x00, 0x00, 0x49,
            0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82         # IEND chunk
        )
        [System.IO.File]::WriteAllBytes($TestImagePath, $pngData)
    }
    
    # Store original functions to restore after mocking
    $Global:OriginalFunctions = @{}
}

Describe "OCRFunctions - Utility Functions" {
    
    Context "Get-CurrentProcessId" {
        It "Should return the current process ID" {
            $result = Get-CurrentProcessId
            $result | Should -Be $PID
            $result | Should -BeOfType [int]
        }
        
        It "Should handle verbose output correctly" {
            $result = Get-CurrentProcessId -Verbose
            $result | Should -Be $PID
        }
    }
    
    Context "Get-ParentProcessId" {
        It "Should return parent process ID for valid process" {
            $currentPID = $PID
            $result = Get-ParentProcessId -ProcessId $currentPID
            $result | Should -BeOfType [int]
            $result | Should -BeGreaterThan 0
        }
        
        It "Should throw for invalid process ID" {
            { Get-ParentProcessId -ProcessId 999999 } | Should -Throw
        }
        
        It "Should validate process ID range" {
            { Get-ParentProcessId -ProcessId 0 } | Should -Throw
            { Get-ParentProcessId -ProcessId -1 } | Should -Throw
        }
    }
    
    Context "Find-MSEdgeWebView2Process" {
        It "Should accept valid parent process ID" {
            $result = Find-MSEdgeWebView2Process -ParentProcessId $PID
            $result | Should -BeOfType [array]
        }
        
        It "Should validate parent process ID range" {
            { Find-MSEdgeWebView2Process -ParentProcessId 0 } | Should -Throw
            { Find-MSEdgeWebView2Process -ParentProcessId -1 } | Should -Throw
        }
        
        It "Should handle non-existent parent process gracefully" {
            $result = Find-MSEdgeWebView2Process -ParentProcessId 999999 -ErrorAction SilentlyContinue
            # Should not throw, just return empty array
            $result | Should -BeOfType [array]
        }
    }
}

Describe "OCRFunctions - Window Management" {
    
    Context "Find-WindowByTitle" {
        It "Should return array for any title search" {
            $result = Find-WindowByTitle -Title "NonExistentWindow"
            $result | Should -BeOfType [array]
        }
        
        It "Should find windows with partial title match" {
            # Look for PowerShell windows which should exist during testing
            $result = Find-WindowByTitle -Title "PowerShell"
            $result | Should -BeOfType [array]
        }
        
        It "Should handle empty title gracefully" {
            $result = Find-WindowByTitle -Title ""
            $result | Should -BeOfType [array]
        }
        
        It "Should handle null title parameter" {
            $result = Find-WindowByTitle -Title $null
            $result | Should -BeOfType [array]
        }
    }
    
    Context "Activate-Window" {
        It "Should handle invalid window handle gracefully" {
            $invalidHandle = [IntPtr]::Zero
            $result = Activate-Window -WindowHandle $invalidHandle
            $result | Should -Be $false
        }
        
        It "Should return boolean result" {
            $invalidHandle = [IntPtr]::Zero
            $result = Activate-Window -WindowHandle $invalidHandle
            $result | Should -BeOfType [bool]
        }
    }
    
    Context "Show-Process" {
        It "Should handle process with MainWindowHandle" {
            $currentProcess = Get-Process -Id $PID
            if ($currentProcess.MainWindowHandle -ne [IntPtr]::Zero) {
                { Show-Process -Process $currentProcess } | Should -Not -Throw
            } else {
                # Skip if no main window handle
                Set-ItResult -Skipped -Because "Process has no main window handle"
            }
        }
        
        It "Should accept Maximize switch parameter" {
            $currentProcess = Get-Process -Id $PID
            if ($currentProcess.MainWindowHandle -ne [IntPtr]::Zero) {
                { Show-Process -Process $currentProcess -Maximize } | Should -Not -Throw
            } else {
                Set-ItResult -Skipped -Because "Process has no main window handle"
            }
        }
    }
}

Describe "OCRFunctions - Input and Screenshot Functions" {
    
    Context "Click-Coordinates" {
        It "Should accept valid coordinates" {
            # Test with screen center coordinates to avoid clicking outside screen
            { Click-Coordinates -X 100 -Y 100 } | Should -Not -Throw
        }
        
        It "Should require mandatory X parameter" {
            { Click-Coordinates -Y 100 } | Should -Throw
        }
        
        It "Should require mandatory Y parameter" {
            { Click-Coordinates -X 100 } | Should -Throw
        }
        
        It "Should handle negative coordinates" {
            # Should not throw, but may not work as expected
            { Click-Coordinates -X -10 -Y -10 } | Should -Not -Throw
        }
    }
    
    Context "Take-Screenshot" {
        BeforeEach {
            $TestScreenshotPath = Join-Path $TestDataPath "test-screenshot.png"
            if (Test-Path $TestScreenshotPath) {
                Remove-Item $TestScreenshotPath -Force
            }
        }
        
        It "Should create screenshot file" {
            $TestScreenshotPath = Join-Path $TestDataPath "test-screenshot.png"
            $result = Take-Screenshot -Path $TestScreenshotPath
            $result | Should -Be $true
            Test-Path $TestScreenshotPath | Should -Be $true
        }
        
        It "Should return boolean result" {
            $TestScreenshotPath = Join-Path $TestDataPath "test-screenshot2.png"
            $result = Take-Screenshot -Path $TestScreenshotPath
            $result | Should -BeOfType [bool]
        }
        
        It "Should handle invalid path gracefully" {
            $invalidPath = "Z:\Invalid\Path\screenshot.png"
            $result = Take-Screenshot -Path $invalidPath
            $result | Should -Be $false
        }
        
        AfterEach {
            # Clean up test files
            $TestScreenshotPath = Join-Path $TestDataPath "test-screenshot.png"
            $TestScreenshotPath2 = Join-Path $TestDataPath "test-screenshot2.png"
            if (Test-Path $TestScreenshotPath) { Remove-Item $TestScreenshotPath -Force -ErrorAction SilentlyContinue }
            if (Test-Path $TestScreenshotPath2) { Remove-Item $TestScreenshotPath2 -Force -ErrorAction SilentlyContinue }
        }
    }
    
    Context "Send-Text" {
        It "Should accept text parameter" {
            # Note: This will actually type text, so use minimal test
            { Send-Text -Text "test" } | Should -Not -Throw
        }
        
        It "Should accept delay parameters" {
            { Send-Text -Text "a" -DelayBetweenKeys 1 -DelayAfterClear 1 } | Should -Not -Throw
        }
        
        It "Should accept ClearFirst switch" {
            { Send-Text -Text "a" -ClearFirst } | Should -Not -Throw
        }
        
        It "Should require mandatory Text parameter" {
            { Send-Text -DelayBetweenKeys 10 } | Should -Throw
        }
    }
}

Describe "OCRFunctions - OCR Functions" {
    
    Context "Find-TextInImageUsingWindowsOCR" {
        BeforeAll {
            $TestImagePath = Join-Path $TestDataPath "test-image.png"
        }
        
        It "Should require mandatory ImagePath parameter" {
            { Find-TextInImageUsingWindowsOCR -SearchText "test" } | Should -Throw
        }
        
        It "Should require mandatory SearchText parameter" {
            { Find-TextInImageUsingWindowsOCR -ImagePath $TestImagePath } | Should -Throw
        }
        
        It "Should throw for non-existent image file" {
            $nonExistentPath = "C:\NonExistent\image.png"
            { Find-TextInImageUsingWindowsOCR -ImagePath $nonExistentPath -SearchText "test" } | Should -Throw
        }
        
        It "Should return array result for valid image" {
            if (Test-Path $TestImagePath) {
                $result = Find-TextInImageUsingWindowsOCR -ImagePath $TestImagePath -SearchText "test" -ErrorAction SilentlyContinue
                $result | Should -BeOfType [array]
            } else {
                Set-ItResult -Skipped -Because "Test image not available"
            }
        }
        
        It "Should accept CaseSensitive switch" {
            if (Test-Path $TestImagePath) {
                $result = Find-TextInImageUsingWindowsOCR -ImagePath $TestImagePath -SearchText "test" -CaseSensitive -ErrorAction SilentlyContinue
                $result | Should -BeOfType [array]
            } else {
                Set-ItResult -Skipped -Because "Test image not available"
            }
        }
        
        It "Should accept OutputImagePath parameter" {
            if (Test-Path $TestImagePath) {
                $outputPath = Join-Path $TestDataPath "output-test.png"
                $result = Find-TextInImageUsingWindowsOCR -ImagePath $TestImagePath -SearchText "test" -OutputImagePath $outputPath -ErrorAction SilentlyContinue
                $result | Should -BeOfType [array]
                # Clean up
                if (Test-Path $outputPath) { Remove-Item $outputPath -Force -ErrorAction SilentlyContinue }
            } else {
                Set-ItResult -Skipped -Because "Test image not available"
            }
        }
        
        It "Should accept MatchIndex parameter" {
            if (Test-Path $TestImagePath) {
                $result = Find-TextInImageUsingWindowsOCR -ImagePath $TestImagePath -SearchText "test" -MatchIndex 0 -ErrorAction SilentlyContinue
                $result | Should -BeOfType [array]
            } else {
                Set-ItResult -Skipped -Because "Test image not available"
            }
        }
        
        It "Should handle phrase search (text with spaces)" {
            if (Test-Path $TestImagePath) {
                $result = Find-TextInImageUsingWindowsOCR -ImagePath $TestImagePath -SearchText "hello world" -ErrorAction SilentlyContinue
                $result | Should -BeOfType [array]
            } else {
                Set-ItResult -Skipped -Because "Test image not available"
            }
        }
    }
    
    Context "Get-AllTextFromImage" {
        BeforeAll {
            $TestImagePath = Join-Path $TestDataPath "test-image.png"
        }
        
        It "Should require mandatory ImagePath parameter" {
            { Get-AllTextFromImage } | Should -Throw
        }
        
        It "Should return string result for valid image" {
            if (Test-Path $TestImagePath) {
                $result = Get-AllTextFromImage -ImagePath $TestImagePath -ErrorAction SilentlyContinue
                $result | Should -BeOfType [string]
            } else {
                Set-ItResult -Skipped -Because "Test image not available"
            }
        }
        
        It "Should handle non-existent image file" {
            $nonExistentPath = "C:\NonExistent\image.png"
            $result = Get-AllTextFromImage -ImagePath $nonExistentPath -ErrorAction SilentlyContinue
            $result | Should -BeOfType [string]
            $result | Should -BeExactly ""
        }
    }
}

Describe "OCRFunctions - Integration Tests" {
    
    Context "End-to-End OCR Workflow" {
        It "Should perform complete OCR workflow" -Skip:(-not (Test-Path (Join-Path $TestDataPath "test-image.png"))) {
            $TestImagePath = Join-Path $TestDataPath "test-image.png"
            
            # Take screenshot
            $screenshotPath = Join-Path $TestDataPath "integration-test-screenshot.png"
            $screenshotResult = Take-Screenshot -Path $screenshotPath
            $screenshotResult | Should -Be $true
            
            # Perform OCR on screenshot
            $ocrResult = Get-AllTextFromImage -ImagePath $screenshotPath -ErrorAction SilentlyContinue
            $ocrResult | Should -BeOfType [string]
            
            # Clean up
            if (Test-Path $screenshotPath) {
                Remove-Item $screenshotPath -Force -ErrorAction SilentlyContinue
            }
        }
    }
    
    Context "Error Handling" {
        It "Should handle OCR engine unavailability gracefully" {
            # This tests the error handling when OCR engine is not available
            Mock -CommandName "[Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages" -MockWith { return $null } -ModuleName $null
            
            $TestImagePath = Join-Path $TestDataPath "test-image.png"
            if (Test-Path $TestImagePath) {
                { Find-TextInImageUsingWindowsOCR -ImagePath $TestImagePath -SearchText "test" } | Should -Throw
            }
        }
    }
}

AfterAll {
    # Clean up test data
    $TestDataPath = Join-Path $PSScriptRoot "TestData"
    if (Test-Path $TestDataPath) {
        Get-ChildItem $TestDataPath -Filter "test-*" | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem $TestDataPath -Filter "integration-*" | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem $TestDataPath -Filter "output-*" | Remove-Item -Force -ErrorAction SilentlyContinue
    }
}