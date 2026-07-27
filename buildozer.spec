[app]
title = AIS140 Sender
package.name = ais140sender
package.domain = org.yourcompany
source.dir = .
source.include_exts = py,png,jpg,kv,atlas
version = 1.0
[app]
requirements=python3==3.12,kivy==2.3.0,requests,openpyxl,certifi,urllib3,charset_normalizer,idna
orientation = portrait
fullscreen = 0

android.permissions = INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.archs = arm64-v8a,armeabi-v7a

[buildozer]
log_level = 2
warn_on_root = 1
