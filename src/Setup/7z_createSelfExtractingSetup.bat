del b2xtranslator_setup.7z
del b2xtranslator_setup.exe
7z_extras\7zr a b2xtranslator_setup.7z Win32Installer\Release\*.*
copy /b 7z_extras\7zSD.sfx + 7z_config.txt + b2xtranslator_setup.7z b2xtranslator_setup.exe
del b2xtranslator_setup.7z