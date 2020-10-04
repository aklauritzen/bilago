rmdir /S /Q dist
rmdir /S /Q build

pyinstaller bilago.py --noconfirm --log-level=WARN^
    --onefile ^
    --add-data="C:/Users/Anders/Dropbox/Python/Projects/bilago/static/images/gear_loader.png;static/images" ^
    --add-data="C:/Users/Anders/Dropbox/Python/Projects/bilago/static/images/header.png;static/images" ^
    --hidden-import pkg_resources.py2_warn ^
    --icon="bilagoicon.ico"