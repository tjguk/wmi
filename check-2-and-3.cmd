cd %TEMP%
if EXIST wmi-release rmdir/s/q wmi-release
mkdir wmi-release
cd wmi-release

git clone git@github.com:tjguk/wmi.git wmi2
pushd wmi2
py -2 -m virtualenv .venv
.venv\scripts\pip install -e .[all]
popd

git clone git@github.com:tjguk/wmi.git wmi3
pushd wmi3
py -3 -m venv .venv
.venv\scripts\pip install -e .[all]
popd

PAUSE