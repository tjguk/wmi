cd %TEMP%
if EXIST wmi-release rmdir/s/q wmi-release
mkdir wmi-release
cd wmi-release

git clone git@github.com:tjguk/wmi.git
cd wmi
py -3 -mvenv .venv
.venv\scripts\pip install -e .[all]
.venv\scripts\python setup.py sdist bdist_wheel --universal
.venv\scripts\python setup.py sdist
.venv\scripts\twine check dist/*
.venv\scripts\twine upload dist/*

PAUSE