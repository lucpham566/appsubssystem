import zipfile

PROXY_HOST = "173.239.219.126"
PROXY_PORT = "6035"
PROXY_USER = "wlhkebmr"
PROXY_PASS = "xcxbfhgvisbs"

manifest_json = """{
    "version": "1.0.0",
    "manifest_version": 2,
    "name": "Chrome Proxy",
    "permissions": ["proxy", "tabs", "unlimitedStorage", "storage", "<all_urls>", "webRequest", "webRequestBlocking"],
    "background": {
        "scripts": ["background.js"]
    },
    "minimum_chrome_version":"22.0.0"
}"""

background_js = f"""
var config = {{
    mode: "fixed_servers",
    rules: {{
        singleProxy: {{
            scheme: "http",
            host: "{PROXY_HOST}",
            port: parseInt("{PROXY_PORT}")
        }},
        bypassList: ["localhost"]
    }}
}};

chrome.proxy.settings.set({{value: config, scope: "regular"}}, function() {{}});

chrome.webRequest.onAuthRequired.addListener(
    function(details) {{
        return {{
            authCredentials: {{
                username: "{PROXY_USER}",
                password: "{PROXY_PASS}"
            }}
        }};
    }},
    {{urls: ["<all_urls>"]}},
    ["blocking"]
);
"""

plugin_file = "proxy_auth_plugin.zip"

with zipfile.ZipFile(plugin_file, "w") as zp:
    zp.writestr("manifest.json", manifest_json)
    zp.writestr("background.js", background_js)

print(f"✅ Tạo file {plugin_file} thành công!")
