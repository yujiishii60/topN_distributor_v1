Write-Host "[info] setup (GUI python): C:\DevTools\python311_tk"
$PY = "C:\DevTools\python311_tk\python.exe"
& $PY -m pip install --upgrade pip
& $PY -m pip install -r requirements.txt
& $PY -c "import tkinter; import tkcalendar; import babel; print('smoke ok')"
Write-Host "[ok] deps installed"
