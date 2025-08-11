# OCRFunctions.ps1

# Comprehensive OCR and Window Management Functions for PowerShell

# Add required assemblies
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# Load Windows Runtime assemblies for OCR
Add-Type -AssemblyName System.Runtime.WindowsRuntime

$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]

# Load Windows Runtime assemblies properly
[void][Windows.ApplicationModel.Core.CoreApplication,Windows.ApplicationModel,ContentType=WindowsRuntime]
[void][Windows.Storage.Streams.RandomAccessStreamReference,Windows.Storage.Streams,ContentType=WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics,ContentType=WindowsRuntime]
[void][Windows.Media.Ocr.OcrEngine,Windows.Media.Ocr,ContentType=WindowsRuntime]

# Add Win32 API functions for window enumeration and control
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    [DllImport("user32.dll")]
    public static extern bool IsWindowEnabled(IntPtr hWnd);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    public const uint MOUSEEVENTF_LEFTDOWN = 0x02;
    public const uint MOUSEEVENTF_LEFTUP = 0x04;
    public const uint GW_OWNER = 4;
}

"@

# Functions
Function Show-Process($Process, [Switch]$Maximize)
{
  $sig = '
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern int SetForegroundWindow(IntPtr hwnd);
  '

  if ($Maximize) { $Mode = 3 } else { $Mode = 4 }
  $type = Add-Type -MemberDefinition $sig -Name WindowAPI -PassThru
  $hwnd = $process.MainWindowHandle
  $null = $type::ShowWindowAsync($hwnd, $Mode)
  $null = $type::SetForegroundWindow($hwnd)
}

Function Click-Coordinates {
    param(
        [Parameter(Mandatory=$true)]
        [int]$X,

        [Parameter(Mandatory=$true)]
        [int]$Y
    )

    # Set cursor position and click
    [Win32]::SetCursorPos($X, $Y)
    Start-Sleep -Milliseconds 50
    [Win32]::mouse_event([Win32]::MOUSEEVENTF_LEFTDOWN, $X, $Y, 0, 0)
    [Win32]::mouse_event([Win32]::MOUSEEVENTF_LEFTUP, $X, $Y, 0, 0)
}

Function Activate-Window {
    param([IntPtr]$WindowHandle)

    try {
        [Win32]::ShowWindow($WindowHandle, [Win32]::SW_RESTORE)
        [Win32]::SetForegroundWindow($WindowHandle)
        Start-Sleep -Milliseconds 500  # Wait for window to become active
        return $true
    }
    catch {
        Write-Error "Failed to activate window: $($_.Exception.Message)"
        return $false
    }
}

Function Get-CurrentProcessId {
    <#
    .SYNOPSIS
        Gets the current PowerShell process ID
    .DESCRIPTION
        Returns the process ID of the current PowerShell session
    #>
    [CmdletBinding()]
    Param()

    try {
        Write-Verbose "Getting current process ID"
        return $PID
    } catch {
        Write-Error "Failed to get current process ID: $($_.Exception.Message)"
        throw
    }
}

Function Get-ParentProcessId {
    <#
    .SYNOPSIS
        Gets the parent process ID for a given process
    .PARAMETER ProcessId
        The process ID to find the parent for
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$ProcessId
    )

    try {
        Write-Verbose "Getting parent process ID for PID: $ProcessId"
        $parentPID = (Get-WmiObject -Class Win32_Process -Filter "ProcessId=$ProcessId").ParentProcessId
        Write-Verbose "Parent process ID: $parentPID"
        return $parentPID
    } catch {
        Write-Error "Failed to get parent process ID: $($_.Exception.Message)"
        throw
    }
}

Function Find-MSEdgeWebView2Process {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$ParentProcessId
    )

    try {
        Write-Verbose "Looking for MSEdgeWebView2 processes with parent PID: $ParentProcessId"

        $webviewProcesses = @()
        $childProcesses = Get-WmiObject -Class Win32_Process | Where-Object { ($_.ParentProcessId -eq $ParentProcessId) -And ($_.MainWindowHandle -ne 0) }

        foreach ($child in $childProcesses) {
            $process = Get-Process -Id $child.ProcessId -ErrorAction SilentlyContinue
            if ($process -and $process.ProcessName -like "*MSEdgeWebView2*") {
                $webviewProcesses += $process
            }
        }

        Write-Verbose "Found $($webviewProcesses.Count) MSEdgeWebView2 processes"
        return $webviewProcesses
    } catch {
        Write-Error "Failed to find MSEdgeWebView2 processes: $($_.Exception.Message)"
        throw
    }
}

Function Find-WindowByTitle {
    param([string]$Title)

    $script:foundWindows = @()

    $callback = {
        param($hWnd, $lParam)

        if ([Win32]::IsWindowVisible($hWnd)) {
            $windowText = New-Object System.Text.StringBuilder 256
            [Win32]::GetWindowText($hWnd, $windowText, 256)
            $windowTitle = $windowText.ToString()

            if ($windowTitle -like "*$Title*") {
                $script:foundWindows += $hWnd
            }
        }
        return $true  # Continue enumeration
    }

    # Call EnumWindows - return value here is BOOL success which you should ignore
    [void][Win32]::EnumWindows($callback, [IntPtr]::Zero)

    # Return found windows
    return $foundWindows
}

Function Find-TextInImageUsingWindowsOCR {

    param(
        [Parameter(Mandatory=$true)]
        [string]$ImagePath,
        [Parameter(Mandatory=$true)]
        [string]$SearchText,
        [switch]$CaseSensitive,
        [string]$OutputImagePath = $null,
        [System.Drawing.Color]$BoundingBoxColor = [System.Drawing.Color]::Red,
        [int]$BoundingBoxThickness = 2,
        [int]$MatchIndex = -1
    )

    try {
        # Validate image file exists
        if (-not (Test-Path $ImagePath)) {
            throw "Image file not found: $ImagePath"
        }

        Write-Host "Loading image for OCR analysis: $ImagePath"

        # Load Windows Runtime types
        Add-Type -AssemblyName System.Runtime.WindowsRuntime
        [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime] | Out-Null
        [Windows.Media.Ocr.OcrEngine,Windows.Media.Ocr,ContentType=WindowsRuntime] | Out-Null
        [Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime] | Out-Null

        # Get AsTask method for async operations
        $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() |
            Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]

        # Get file as StorageFile
        $getFileTask = [Windows.Storage.StorageFile]::GetFileFromPathAsync($ImagePath)
        $makeGenericAsTask = $asTaskGeneric.MakeGenericMethod([Windows.Storage.StorageFile])
        $fileTask = $makeGenericAsTask.Invoke($null, @($getFileTask))
        $storageFile = $fileTask.Result

        # Open the file stream
        $openReadTask = $storageFile.OpenReadAsync()
        $makeGenericAsTask2 = $asTaskGeneric.MakeGenericMethod([Windows.Storage.Streams.IRandomAccessStreamWithContentType])
        $streamTask = $makeGenericAsTask2.Invoke($null, @($openReadTask))
        $fileStream = $streamTask.Result

        # Create bitmap decoder
        $createDecoderTask = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($fileStream)
        $makeGenericAsTask3 = $asTaskGeneric.MakeGenericMethod([Windows.Graphics.Imaging.BitmapDecoder])
        $decoderTask = $makeGenericAsTask3.Invoke($null, @($createDecoderTask))
        $decoder = $decoderTask.Result

        # Get software bitmap
        $getSoftwareBitmapTask = $decoder.GetSoftwareBitmapAsync()
        $makeGenericAsTask4 = $asTaskGeneric.MakeGenericMethod([Windows.Graphics.Imaging.SoftwareBitmap])
        $bitmapTask = $makeGenericAsTask4.Invoke($null, @($getSoftwareBitmapTask))
        $softwareBitmap = $bitmapTask.Result

        # Get OCR engine for current language
        $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
        if ($null -eq $ocrEngine) {
            throw "No OCR engine available for current language. Available languages: $([Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages -join ', ')"
        }

        Write-Host "Running OCR analysis with language: $($ocrEngine.RecognizerLanguage.DisplayName)"

        # Perform OCR
        $recognizeTask = $ocrEngine.RecognizeAsync($softwareBitmap)
        $makeGenericAsTask5 = $asTaskGeneric.MakeGenericMethod([Windows.Media.Ocr.OcrResult])
        $ocrTask = $makeGenericAsTask5.Invoke($null, @($recognizeTask))
        $ocrResult = $ocrTask.Result

        Write-Host "OCR completed. Text angle: $($ocrResult.TextAngle)"

        # Find matching text (words and phrases)

        $matches = @()

        # If search text contains spaces, treat it as a phrase search

        if ($SearchText.Contains(" ")) {

            Write-Host "Searching for phrase: '$SearchText'"

            # Split the search phrase into words

            $searchWords = $SearchText.Split(' ', [StringSplitOptions]::RemoveEmptyEntries)

            # Create a flat list of all words with their positions

            $allWords = @()

            $lineIndex = 0

            foreach ($line in $ocrResult.Lines) {

                $wordIndex = 0

                foreach ($word in $line.Words) {

                    $allWords += [PSCustomObject]@{

                        Text = $word.Text

                        BoundingRect = $word.BoundingRect

                        LineIndex = $lineIndex

                        WordIndex = $wordIndex

                    }

                    $wordIndex++

                }

                $lineIndex++

            }

            # Look for consecutive word matches that form the phrase

            for ($i = 0; $i -le ($allWords.Count - $searchWords.Count); $i++) {

                $foundPhrase = $true

                $phraseWords = @()

                for ($j = 0; $j -lt $searchWords.Count; $j++) {

                    $currentWord = $allWords[$i + $j]

                    $searchWord = $searchWords[$j]

                    $wordMatch = if ($CaseSensitive) {

                        $currentWord.Text -ceq $searchWord

                    } else {

                        $currentWord.Text -ieq $searchWord

                    }

                    if (-not $wordMatch) {

                        $foundPhrase = $false

                        break

                    }

                    $phraseWords += $currentWord

                }

                if ($foundPhrase) {

                    # Calculate bounding box that encompasses all words in the phrase

                    $leftValues = $phraseWords | ForEach-Object { $_.BoundingRect.X }

                    $topValues = $phraseWords | ForEach-Object { $_.BoundingRect.Y }

                    $rightValues = $phraseWords | ForEach-Object { $_.BoundingRect.X + $_.BoundingRect.Width }

                    $bottomValues = $phraseWords | ForEach-Object { $_.BoundingRect.Y + $_.BoundingRect.Height }

                    $left = ($leftValues | Measure-Object -Minimum).Minimum

                    $top = ($topValues | Measure-Object -Minimum).Minimum

                    $right = ($rightValues | Measure-Object -Maximum).Maximum

                    $bottom = ($bottomValues | Measure-Object -Maximum).Maximum

                    $phraseText = ($phraseWords | ForEach-Object { $_.Text }) -join " "

                    $matches += [PSCustomObject]@{

                        Text = $phraseText

                        Left = [int]$left

                        Top = [int]$top

                        Width = [int]($right - $left)

                        Height = [int]($bottom - $top)

                        CenterX = [int]($left + (($right - $left) / 2))

                        CenterY = [int]($top + (($bottom - $top) / 2))

                        Confidence = [math]::Round(($right - $left) * ($bottom - $top), 2)

                        Type = "Phrase"

                        WordCount = $phraseWords.Count

                    }

                    # Skip ahead to avoid overlapping matches

                    $i += $searchWords.Count - 1

                }

            }

        }

        else {

            Write-Host "Searching for single word: '$SearchText'"

            # Single word search (original logic)

            $searchPattern = if ($CaseSensitive) { "*$SearchText*" } else { "*$SearchText*" }

            foreach ($line in $ocrResult.Lines) {

                foreach ($word in $line.Words) {

                    $textMatch = if ($CaseSensitive) {

                        $word.Text -clike $searchPattern

                    } else {

                        $word.Text -like $searchPattern

                    }

                    if ($textMatch) {

                        $boundingRect = $word.BoundingRect

                        $matches += [PSCustomObject]@{

                            Text = $word.Text

                            Left = [int]$boundingRect.X

                            Top = [int]$boundingRect.Y

                            Width = [int]$boundingRect.Width

                            Height = [int]$boundingRect.Height

                            CenterX = [int]($boundingRect.X + ($boundingRect.Width / 2))

                            CenterY = [int]($boundingRect.Y + ($boundingRect.Height / 2))

                            Confidence = [math]::Round($word.BoundingRect.Width * $word.BoundingRect.Height, 2)

                            Type = "Word"

                            WordCount = 1

                        }

                    }

                }

            }

        }

        Write-Host "Found $($matches.Count) matches for '$SearchText'"

        # Filter to specific match index if requested

        if ($MatchIndex -ge 0) {

            if ($MatchIndex -lt $matches.Count) {

                Write-Host "Returning match #$MatchIndex : '$($matches[$MatchIndex].Text)'"

                $matches = @($matches[$MatchIndex])

            } else {

                Write-Warning "Match index $MatchIndex not found. Only $($matches.Count) matches available (0-$($matches.Count - 1))"

                $matches = @()

            }

        } else {

            Write-Host "Returning all matches"

        }

        # Create annotated image if requested

        if ($OutputImagePath -and $matches.Count -gt 0) {

            try {

                Write-Host "Creating annotated image with bounding boxes..."

                Add-Type -AssemblyName System.Drawing

                # Load the original image

                $originalImage = [System.Drawing.Image]::FromFile($ImagePath)

                $graphics = [System.Drawing.Graphics]::FromImage($originalImage)

                $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

                # Create pen for drawing bounding boxes

                $pen = New-Object System.Drawing.Pen($BoundingBoxColor, $BoundingBoxThickness)

                # Draw bounding boxes around found text

                foreach ($match in $matches) {

                    $rect = New-Object System.Drawing.Rectangle($match.Left, $match.Top, $match.Width, $match.Height)

                    $graphics.DrawRectangle($pen, $rect)

                    # Optionally add text label above the bounding box

                    $font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)

                    $brush = New-Object System.Drawing.SolidBrush($BoundingBoxColor)

                    $labelY = [Math]::Max(0, $match.Top - 15)

                    $graphics.DrawString($match.Text, $font, $brush, $match.Left, $labelY)

                    $font.Dispose()

                    $brush.Dispose()

                }

                # Save the annotated image

                $originalImage.Save($OutputImagePath)

                # Cleanup graphics resources

                $pen.Dispose()

                $graphics.Dispose()

                $originalImage.Dispose()

                Write-Host "Annotated image saved to: $OutputImagePath" -ForegroundColor Green

            }

            catch {

                Write-Warning "Failed to create annotated image: $($_.Exception.Message)"

            }

        }

        # Cleanup resources

        if ($fileStream) { $fileStream.Dispose() }

        if ($softwareBitmap) { $softwareBitmap.Dispose() }

        return $matches

    }

    catch {

        Write-Error "Windows OCR processing failed: $($_.Exception.Message)"

        Write-Host "Error details: $($_.Exception.ToString())" -ForegroundColor Red

        return @()

    }

}

function Get-AllTextFromImage {

    param(

        [Parameter(Mandatory=$true)]

        [string]$ImagePath

    )

    try {

        # Similar setup as above but return all text

        Add-Type -AssemblyName System.Runtime.WindowsRuntime

        [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime] | Out-Null

        [Windows.Media.Ocr.OcrEngine,Windows.Media.Ocr,ContentType=WindowsRuntime] | Out-Null

        [Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime] | Out-Null

        $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() |

            Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]

        # Process image (same as above)

        $getFileTask = [Windows.Storage.StorageFile]::GetFileFromPathAsync($ImagePath)

        $makeGenericAsTask = $asTaskGeneric.MakeGenericMethod([Windows.Storage.StorageFile])

        $fileTask = $makeGenericAsTask.Invoke($null, @($getFileTask))

        $storageFile = $fileTask.Result

        $openReadTask = $storageFile.OpenReadAsync()

        $makeGenericAsTask2 = $asTaskGeneric.MakeGenericMethod([Windows.Storage.Streams.IRandomAccessStreamWithContentType])

        $streamTask = $makeGenericAsTask2.Invoke($null, @($openReadTask))

        $fileStream = $streamTask.Result

        $createDecoderTask = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($fileStream)

        $makeGenericAsTask3 = $asTaskGeneric.MakeGenericMethod([Windows.Graphics.Imaging.BitmapDecoder])

        $decoderTask = $makeGenericAsTask3.Invoke($null, @($createDecoderTask))

        $decoder = $decoderTask.Result

        $getSoftwareBitmapTask = $decoder.GetSoftwareBitmapAsync()

        $makeGenericAsTask4 = $asTaskGeneric.MakeGenericMethod([Windows.Graphics.Imaging.SoftwareBitmap])

        $bitmapTask = $makeGenericAsTask4.Invoke($null, @($getSoftwareBitmapTask))

        $softwareBitmap = $bitmapTask.Result

        $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()

        $recognizeTask = $ocrEngine.RecognizeAsync($softwareBitmap)

        $makeGenericAsTask5 = $asTaskGeneric.MakeGenericMethod([Windows.Media.Ocr.OcrResult])

        $ocrTask = $makeGenericAsTask5.Invoke($null, @($recognizeTask))

        $ocrResult = $ocrTask.Result

        # Return all text as a single string

        $allText = $ocrResult.Text

        # Cleanup

        if ($fileStream) { $fileStream.Dispose() }

        if ($softwareBitmap) { $softwareBitmap.Dispose() }

        return $allText

    }

    catch {

        Write-Error "Failed to extract text: $($_.Exception.Message)"

        return ""

    }

}

function Take-Screenshot {

    param([string]$Path)

    try {

        $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

        $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height

        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)

        $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)

        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)

        $graphics.Dispose()

        $bitmap.Dispose()

        Write-Host "Screenshot saved to: $Path"

        return $true

    }

    catch {

        Write-Error "Failed to take screenshot: $($_.Exception.Message)"

        return $false

    }

}

# Type text function using Windows SendInput API

function Send-Text {

    param(

        [Parameter(Mandatory=$true)]

        [string]$Text,

        [int]$DelayBetweenKeys = 10,

        [switch]$ClearFirst,

        [int]$DelayAfterClear = 100

    )

    try {

        Add-Type -TypeDefinition @"

        using System;

        using System.Runtime.InteropServices;

        using System.Windows.Forms;

        public class KeyboardInput {

            [DllImport("user32.dll", SetLastError = true)]

            public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

            [DllImport("user32.dll")]

            public static extern short VkKeyScan(char ch);

            [StructLayout(LayoutKind.Sequential)]

            public struct INPUT {

                public uint type;

                public INPUTUNION u;

            }

            [StructLayout(LayoutKind.Explicit)]

            public struct INPUTUNION {

                [FieldOffset(0)]

                public KEYBDINPUT ki;

            }

            [StructLayout(LayoutKind.Sequential)]

            public struct KEYBDINPUT {

                public ushort wVk;

                public ushort wScan;

                public uint dwFlags;

                public uint time;

                public IntPtr dwExtraInfo;

            }

            public const uint INPUT_KEYBOARD = 1;

            public const uint KEYEVENTF_KEYUP = 0x0002;

            public const uint KEYEVENTF_UNICODE = 0x0004;

        }

"@ -ReferencedAssemblies System.Windows.Forms

        Write-Host "Typing text: '$Text'"

        # Clear existing text if requested

        if ($ClearFirst) {

            Write-Host "Clearing existing text with Ctrl+A, Delete"

            # Ctrl+A to select all

            $ctrlDown = @{

                type = [KeyboardInput]::INPUT_KEYBOARD

                u = @{

                    ki = @{

                        wVk = 0x11  # VK_CONTROL

                        wScan = 0

                        dwFlags = 0

                        time = 0

                        dwExtraInfo = [IntPtr]::Zero

                    }

                }

            }

            $aDown = @{

                type = [KeyboardInput]::INPUT_KEYBOARD

                u = @{

                    ki = @{

                        wVk = 0x41  # VK_A

                        wScan = 0

                        dwFlags = 0

                        time = 0

                        dwExtraInfo = [IntPtr]::Zero

                    }

                }

            }

            $aUp = @{

                type = [KeyboardInput]::INPUT_KEYBOARD

                u = @{

                    ki = @{

                        wVk = 0x41  # VK_A

                        wScan = 0

                        dwFlags = [KeyboardInput]::KEYEVENTF_KEYUP

                        time = 0

                        dwExtraInfo = [IntPtr]::Zero

                    }

                }

            }

            $ctrlUp = @{

                type = [KeyboardInput]::INPUT_KEYBOARD

                u = @{

                    ki = @{

                        wVk = 0x11  # VK_CONTROL

                        wScan = 0

                        dwFlags = [KeyboardInput]::KEYEVENTF_KEYUP

                        time = 0

                        dwExtraInfo = [IntPtr]::Zero

                    }

                }

            }

            # Send Ctrl+A

            $inputs = @($ctrlDown, $aDown, $aUp, $ctrlUp)

            [KeyboardInput]::SendInput($inputs.Length, $inputs, [System.Runtime.InteropServices.Marshal]::SizeOf([KeyboardInput+INPUT]))

            Start-Sleep -Milliseconds 50

            # Send Delete

            $deleteDown = @{

                type = [KeyboardInput]::INPUT_KEYBOARD

                u = @{

                    ki = @{

                        wVk = 0x2E  # VK_DELETE

                        wScan = 0

                        dwFlags = 0

                        time = 0

                        dwExtraInfo = [IntPtr]::Zero

                    }

                }

            }

            $deleteUp = @{

                type = [KeyboardInput]::INPUT_KEYBOARD

                u = @{

                    ki = @{

                        wVk = 0x2E  # VK_DELETE

                        wScan = 0

                        dwFlags = [KeyboardInput]::KEYEVENTF_KEYUP

                        time = 0

                        dwExtraInfo = [IntPtr]::Zero

                    }

                }

            }

            $inputs = @($deleteDown, $deleteUp)

            [KeyboardInput]::SendInput($inputs.Length, $inputs, [System.Runtime.InteropServices.Marshal]::SizeOf([KeyboardInput+INPUT]))

            Start-Sleep -Milliseconds $DelayAfterClear

        }

        # Type each character

        foreach ($char in $Text.ToCharArray()) {

            if ($char -eq "`n" -or $char -eq "`r") {

                # Handle Enter key

                $enterDown = @{

                    type = [KeyboardInput]::INPUT_KEYBOARD

                    u = @{

                        ki = @{

                            wVk = 0x0D  # VK_RETURN

                            wScan = 0

                            dwFlags = 0

                            time = 0

                            dwExtraInfo = [IntPtr]::Zero

                        }

                    }

                }

                $enterUp = @{

                    type = [KeyboardInput]::INPUT_KEYBOARD

                    u = @{

                        ki = @{

                            wVk = 0x0D  # VK_RETURN

                            wScan = 0

                            dwFlags = [KeyboardInput]::KEYEVENTF_KEYUP

                            time = 0

                            dwExtraInfo = [IntPtr]::Zero

                        }

                    }

                }

                $inputs = @($enterDown, $enterUp)

                [KeyboardInput]::SendInput($inputs.Length, $inputs, [System.Runtime.InteropServices.Marshal]::SizeOf([KeyboardInput+INPUT]))

            }

            elseif ($char -eq "`t") {

                # Handle Tab key

                $tabDown = @{

                    type = [KeyboardInput]::INPUT_KEYBOARD

                    u = @{

                        ki = @{

                            wVk = 0x09  # VK_TAB

                            wScan = 0

                            dwFlags = 0

                            time = 0

                            dwExtraInfo = [IntPtr]::Zero

                        }

                    }

                }

                $tabUp = @{

                    type = [KeyboardInput]::INPUT_KEYBOARD

                    u = @{

                        ki = @{

                            wVk = 0x09  # VK_TAB

                            wScan = 0

                            dwFlags = [KeyboardInput]::KEYEVENTF_KEYUP

                            time = 0

                            dwExtraInfo = [IntPtr]::Zero

                        }

                    }

                }

                $inputs = @($tabDown, $tabUp)

                [KeyboardInput]::SendInput($inputs.Length, $inputs, [System.Runtime.InteropServices.Marshal]::SizeOf([KeyboardInput+INPUT]))

            }

            else {

                # Handle regular Unicode characters

                $charDown = @{

                    type = [KeyboardInput]::INPUT_KEYBOARD

                    u = @{

                        ki = @{

                            wVk = 0

                            wScan = [uint16][char]$char

                            dwFlags = [KeyboardInput]::KEYEVENTF_UNICODE

                            time = 0

                            dwExtraInfo = [IntPtr]::Zero

                        }

                    }

                }

                $charUp = @{

                    type = [KeyboardInput]::INPUT_KEYBOARD

                    u = @{

                        ki = @{

                            wVk = 0

                            wScan = [uint16][char]$char

                            dwFlags = [KeyboardInput]::KEYEVENTF_UNICODE -bor [KeyboardInput]::KEYEVENTF_KEYUP

                            time = 0

                            dwExtraInfo = [IntPtr]::Zero

                        }

                    }

                }

                $inputs = @($charDown, $charUp)

                [KeyboardInput]::SendInput($inputs.Length, $inputs, [System.Runtime.InteropServices.Marshal]::SizeOf([KeyboardInput+INPUT]))

            }

            if ($DelayBetweenKeys -gt 0) {

                Start-Sleep -Milliseconds $DelayBetweenKeys

            }

        }

        Write-Host "Text typing completed" -ForegroundColor Green

    }

    catch {

        Write-Error "Failed to send text: $($_.Exception.Message)"

    }

}

