7z a -t7z Softscan.7z Office2003PIA\* Office2007PIA\* setup.exe Setup.vbs SpamGrabberSetupSoftScan.msi
copy /b "C:\Program Files\7-Zip\7zSD.sfx" + config.txt + Softscan.7z spamgrabbersoftscan.exe