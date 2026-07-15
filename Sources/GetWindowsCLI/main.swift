import AppKit
import Vision

struct OcrLine {
	let text: String
	let confidence: Float
	let x: CGFloat
	let y: CGFloat
}

let aiAppTitleOcrBundleIdentifiers = Set([
	"com.anthropic.claudefordesktop",
	"com.openai.codex"
])

let ignoredAiAppTitleOcrText = Set([
	"artifacts",
	"chat",
	"chats",
	"code",
	"cowork",
	"customize",
	"dispatch",
	"from calendar",
	"from drive",
	"from gmail",
	"home",
	"learn",
	"new",
	"no chats",
	"open in",
	"plugins",
	"projects",
	"recents",
	"scheduled",
	"search",
	"show more",
	"write"
])

func cleanOcrText(_ text: String) -> String {
	text
		.trimmingCharacters(in: .whitespacesAndNewlines)
		.trimmingCharacters(in: CharacterSet(charactersIn: "•·-–—→←<>‹›+*|"))
		.trimmingCharacters(in: .whitespacesAndNewlines)
}

func normalizedOcrText(_ text: String) -> String {
	text
		.lowercased()
		.components(separatedBy: CharacterSet.alphanumerics.inverted)
		.filter { !$0.isEmpty }
		.joined(separator: " ")
}

func isEmptyTitle(_ title: String) -> Bool {
	let normalizedTitle = title.trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
	return normalizedTitle.isEmpty
}

@available(macOS 10.15, *)
func recognizeTextInWindow(windowID: CGWindowID) -> [OcrLine] {
	guard
		let image = CGWindowListCreateImage(
			.null,
			.optionIncludingWindow,
			windowID,
			[.boundsIgnoreFraming, .bestResolution]
		)
	else {
		return []
	}

	var lines = [OcrLine]()
	let request = VNRecognizeTextRequest { request, _ in
		let observations = request.results as? [VNRecognizedTextObservation] ?? []
		lines = observations.compactMap { observation in
			guard let candidate = observation.topCandidates(1).first else {
				return nil
			}

			let text = cleanOcrText(candidate.string)
			if text.isEmpty || text.count > 160 || candidate.confidence < 0.35 {
				return nil
			}

			return OcrLine(
				text: text,
				confidence: candidate.confidence,
				x: observation.boundingBox.origin.x,
				y: observation.boundingBox.origin.y
			)
		}
	}
	request.recognitionLevel = .accurate
	request.usesLanguageCorrection = true

	let handler = VNImageRequestHandler(cgImage: image, options: [:])
	do {
		try handler.perform([request])
	} catch {
		return []
	}

	return lines.sorted {
		if abs($0.y - $1.y) > 0.02 {
			return $0.y > $1.y
		}
		return $0.x < $1.x
	}
}

func aiAppTitleCandidate(lines: [OcrLine], currentTitle: String, appName: String) -> String? {
	guard isEmptyTitle(currentTitle) else {
		return nil
	}

	for line in lines where line.y >= 0.6 && line.x >= 0.14 {
		let text = cleanOcrText(line.text)
		let normalized = normalizedOcrText(text)

		if text.count < 4 ||
			ignoredAiAppTitleOcrText.contains(normalized) ||
			normalized == appName.lowercased() ||
			normalized.hasPrefix("how can i help you") {
			continue
		}

		return text
	}

	return nil
}

func getActiveBrowserTabURLAppleScriptCommand(_ appId: String) -> String? {
	switch appId {
	case "com.google.Chrome", "com.google.Chrome.beta", "com.google.Chrome.dev", "com.google.Chrome.canary", "com.brave.Browser", "com.brave.Browser.beta", "com.brave.Browser.nightly", "com.microsoft.edgemac", "com.microsoft.edgemac.Beta", "com.microsoft.edgemac.Dev", "com.microsoft.edgemac.Canary", "com.mighty.app", "com.ghostbrowser.gb1", "com.bookry.wavebox", "com.pushplaylabs.sidekick", "com.operasoftware.Opera", "com.operasoftware.OperaNext", "com.operasoftware.OperaDeveloper", "com.vivaldi.Vivaldi", "ru.yandex.desktop.yandex-browser", "com.operasoftware.OperaGX", "ai.perplexity.comet":
		return """
			tell app id \"\(appId)\"
				set window_url to URL of active tab of front window
				set window_name to title of active tab of front window
				set window_mode to mode of front window
				set window_data to window_url & "+++++" & window_name & "+++++" & window_mode
			end tell
			window_data
			"""
	case "com.apple.Safari", "com.apple.SafariTechnologyPreview":
		return """
			tell app id \"\(appId)\"
				set window_url to URL of front document
				set window_name to name of front document
				set window_mode to "normal"
				set window_data to window_url & "+++++" & window_name & "+++++" & window_mode
			end tell
			window_data
			"""
	case "com.kagi.kagimacOS":
		return """
			tell app id \"\(appId)\"
				set window_url to URL of front document
				set window_name to name of front window
				set window_mode to "normal"
				set window_data to window_url & "+++++" & window_name & "+++++" & window_mode
			end tell
			window_data
			"""
	case "company.thebrowser.dia":
		return """
			tell app id \"\(appId)\"
				tell front window
					set window_url to \"\"
					set window_name to \"\"
					repeat with t in tabs
						if isFocused of t is true then
							set window_url to URL of t
							set window_name to title of t
							exit repeat
						end if
					end repeat
					set window_mode to \"normal\"
					set window_data to window_url & \"+++++\" & window_name & \"+++++\" & window_mode
				end tell
			end tell
			window_data
			"""
	case "company.thebrowser.Browser", "com.sigmaos.sigmaos.macos", "com.SigmaOS.SigmaOS":
		return """
			tell app id \"\(appId)\"
				set window_url to URL of active tab of front window
				set window_name to name of active tab of front window
				set window_mode to mode of front window
				set window_data to window_url & "+++++" & window_name & "+++++" & window_mode
			end tell
			window_data
			"""
	default:
		return """
			tell application "System Events"
				tell (first process whose frontmost is true)
					set window_url to ""
					set window_name to value of attribute "AXTitle" of window 1
					set window_mode to "normal"
					set window_data to window_url & "+++++" & window_name & "+++++" & window_mode
				end tell
			end tell
			window_data
			"""
	}
}

func exitWithoutResult() -> Never {
	print("null")
	exit(0)
}

func printOutput(_ output: Any) -> Never {
	guard let string = try? toJson(output) else {
		exitWithoutResult()
	}

	print(string)
	exit(0)
}

func getWindowInformation(window: [String: Any], windowOwnerPID: pid_t) -> [String: Any]? {
	// Skip transparent windows, like with Chrome.
	if (window[kCGWindowAlpha as String] as! Double) == 0 { // Documented to always exist.
		return nil
	}

	let bounds = CGRect(dictionaryRepresentation: window[kCGWindowBounds as String] as! CFDictionary)! // Documented to always exist.

	// Skip tiny windows, like the Chrome link hover statusbar.
	let minWinSize: CGFloat = 50
	if bounds.width < minWinSize || bounds.height < minWinSize {
		return nil
	}

	// This should not fail as we're only dealing with apps, but we guard it just to be safe.
	guard let app = NSRunningApplication(processIdentifier: windowOwnerPID) else {
		return nil
	}

	let appName = window[kCGWindowOwnerName as String] as? String ?? app.bundleIdentifier ?? "<Unknown>"

	let windowTitle = disableScreenRecordingPermission ? "" : window[kCGWindowName as String] as? String ?? ""

	if app.bundleIdentifier == "com.apple.dock" {
		return nil
	}

	var output: [String: Any] = [
		"platform": "macos",
		"title": windowTitle,
		"id": window[kCGWindowNumber as String] as! Int, // Documented to always exist.
		"bounds": [
			"x": bounds.origin.x,
			"y": bounds.origin.y,
			"width": bounds.width,
			"height": bounds.height
		],
		"owner": [
			"name": appName,
			"processId": windowOwnerPID,
			"bundleId": app.bundleIdentifier ?? "", // I don't think this could happen, but we also don't want to crash.
			"path": app.bundleURL?.path ?? "" // I don't think this could happen, but we also don't want to crash.
		],
		"memoryUsage": window[kCGWindowMemoryUsage as String] as? Int ?? 0
	]

	// Run the AppleScript to get the URL if the active window is a compatible browser and accessibility permissions are enabled.
	if
		!disableAccessibilityPermission,
		let bundleIdentifier = app.bundleIdentifier,
		let script = getActiveBrowserTabURLAppleScriptCommand(bundleIdentifier),
		let windowData = runAppleScript(source: script)
	{
		let windowDataArray = windowData.components(separatedBy: "+++++")
		output["url"] = windowDataArray[0]
		output["title"] = windowDataArray[1]
		output["mode"] = windowDataArray[2]
	}

	if
		enableAiAppTitleOcr,
		!enableOpenWindowsList,
		let bundleIdentifier = app.bundleIdentifier,
		aiAppTitleOcrBundleIdentifiers.contains(bundleIdentifier),
		let windowNumber = window[kCGWindowNumber as String] as? Int
	{
		if #available(macOS 10.15, *) {
			let lines = recognizeTextInWindow(windowID: CGWindowID(windowNumber))
			if !lines.isEmpty {
				output["titleOcrText"] = Array(lines.prefix(80).map(\.text))

				let currentTitle = output["title"] as? String ?? ""
				if let title = aiAppTitleCandidate(lines: lines, currentTitle: currentTitle, appName: appName) {
					output["title"] = title
					output["titleSource"] = "ocr"
				}
			}
		}
	}

	return output
}

let disableAccessibilityPermission = CommandLine.arguments.contains("--no-accessibility-permission")
let disableScreenRecordingPermission = CommandLine.arguments.contains("--no-screen-recording-permission")
let enableOpenWindowsList = CommandLine.arguments.contains("--open-windows-list")
let enableAiAppTitleOcr = CommandLine.arguments.contains("--ai-app-title-ocr")

// Show accessibility permission prompt if needed. Required to get the URL of the active tab in browsers.
if !disableAccessibilityPermission {
	if !AXIsProcessTrustedWithOptions(["AXTrustedCheckOptionPrompt": true] as CFDictionary) {
		print("get-windows requires the accessibility permission in “System Settings › Privacy & Security › Accessibility”.")
		exit(1)
	}
}

// Show screen recording permission prompt if needed. Required to get the complete window title.
if
	(!disableScreenRecordingPermission || enableAiAppTitleOcr),
	!hasScreenRecordingPermission()
{
	print("get-windows requires the screen recording permission in “System Settings › Privacy & Security › Screen Recording”.")
	exit(1)
}

guard
	let frontmostAppPID = NSWorkspace.shared.frontmostApplication?.processIdentifier,
	let windows = CGWindowListCopyWindowInfo([.optionOnScreenOnly, .excludeDesktopElements], kCGNullWindowID) as? [[String: Any]]
else {
	exitWithoutResult()
}

var openWindows = [[String: Any]]();

for window in windows {
	let windowOwnerPID = window[kCGWindowOwnerPID as String] as! pid_t // Documented to always exist.
	if !enableOpenWindowsList && windowOwnerPID != frontmostAppPID {
		continue
	}

	guard let windowInformation = getWindowInformation(window: window, windowOwnerPID: windowOwnerPID) else {
		continue
	}

	if !enableOpenWindowsList {
		printOutput(windowInformation)
	} else {
		openWindows.append(windowInformation)
	}
}

if !openWindows.isEmpty {
	printOutput(openWindows)
}

exitWithoutResult()
