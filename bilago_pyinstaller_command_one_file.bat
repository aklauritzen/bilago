rmdir /S /Q dist
rmdir /S /Q build

pyinstaller bilago.py --noconfirm

@REM pyinstaller bilago.py --noconfirm --log-level=WARN ^
@REM     --add-data="C:/Users/Anders/Dropbox/Python/Projects/bilago/static/images/gear_loader.png;static/images" ^
@REM     --add-data="C:/Users/Anders/Dropbox/Python/Projects/bilago/static/images/header.png;static/images" ^
@REM     --hidden-import pkg_resources.py2_warn ^
@REM     --icon="bilagoicon.ico"
