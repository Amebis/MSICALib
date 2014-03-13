@echo off
nmake.exe /ls Clean LANG=En CFG=Debug   PLAT=Win32
nmake.exe /ls Clean LANG=Sl CFG=Debug   PLAT=Win32
nmake.exe /ls Clean LANG=En CFG=Debug   PLAT=x64
nmake.exe /ls Clean LANG=Sl CFG=Debug   PLAT=x64
nmake.exe /ls Clean LANG=En CFG=Release PLAT=Win32
nmake.exe /ls Clean LANG=Sl CFG=Release PLAT=Win32
nmake.exe /ls Clean LANG=En CFG=Release PLAT=x64
nmake.exe /ls Clean LANG=Sl CFG=Release PLAT=x64
