7z a -t7z SpamCop.7z Office2003PIA\* Office2007PIA\* setup.exe Setup.vbs SpamGrabberSetup.msi
copy /b "C:\Program Files\7-Zip\7zSD.sfx" + config.txt + SpamCop.7z spamgrabber.exe