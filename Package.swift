// swift-tools-version:5.9
import PackageDescription

let package = Package(
	name: "GetWindows",
	platforms: [
		.macOS(.v10_13)
	],
	products: [
		.executable(
			name: "get-windows",
			targets: [
				"GetWindowsCLI"
			]
		)
	],
	targets: [
		.executableTarget(
			name: "GetWindowsCLI"
		)
	]
)
