[app]

title = AAHM
package.name = aahm
package.domain = az.edu.aahm

source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,atlas,json,docx
version = 0.1

requirements = python3,kivy,python-docx,lxml

icon.filename = %(source.dir)s/aahm_logo.png

orientation = portrait
fullscreen = 0

android.api = 35
android.minapi = 24
android.ndk = 27c
android.archs = arm64-v8a
android.accept_sdk_license = True

android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

p4a.fork = kivy
p4a.branch = master

[buildozer]

log_level = 2
warn_on_root = 1
