[app]
title = AAHM
package.name = aahm
package.domain = az.edu.aahm
source.dir = .
source.include_exts = py,png,json,docx
version = 0.1
requirements = python3,kivy,python-docx
orientation = portrait
fullscreen = 0

# Android permissions may need adjustment depending on Android version and
# how you choose to import/export files.
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.api = 35
android.minapi = 24
android.archs = arm64-v8a
android.accept_sdk_license = True

[buildozer]
log_level = 2
warn_on_root = 1
