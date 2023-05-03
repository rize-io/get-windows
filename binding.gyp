{
	"targets": [
		{
			"target_name": "<(module_name)",
			"cflags!": [
				"-fno-exceptions"
			],
			"cflags_cc!": [
				"-fno-exceptions"
			],
			"conditions":[
				[
					"OS=='win'",
					{
						"sources": [
							"sources/windows/main.cc",
						],
						'library_dirs': [
							'C:/Program Files (x86)/Windows Kits/10/Lib/10.0.19041.0/um/x64'
						],
						'libraries': [
							'version.lib',
							'Dwmapi.lib',
							'UIAutomationCore.lib'
						],
					},
				],
			],
			"include_dirs": [
				"<!@(node -p \"require('node-addon-api').include\")",
			],
			"defines": [
				"NAPI_VERSION=<(napi_build_version)", "NAPI_DISABLE_CPP_EXCEPTIONS=1",
			],
			'msvs_settings': {
				'VCCLCompilerTool': {
					'ExceptionHandling': 1,
				},
			},
		},
		{
			"target_name": "action_after_build",
			"type": "none",
			"dependencies": [
				"<(module_name)",
			],
			"copies": [
				{
					"files": [
						"<(PRODUCT_DIR)/<(module_name).node",
					],
					"destination": "<(module_path)"
				}
			]
		}
	]
}
