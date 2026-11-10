@echo off
rem Builds the static GitHub Pages version of the site into the docs\ folder (site at /crafty/, enquiry forms open WhatsApp).
rem The normal site in this folder (run.bat / update.bat) is not touched.
python build.py --github-pages
echo.
echo Next: git add -A ^& git commit -m "Update site" ^& git push
pause
