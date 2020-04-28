import os
import re
from setuptools import setup

classifiers = [
    'Development Status :: 5 - Production/Stable',
    'Environment :: Win32 (MS Windows)',
    'Intended Audience :: Developers',
    'Intended Audience :: System Administrators',
    'License :: MIT',
    'Natural Language :: English',
    'Operating System :: Microsoft :: Windows :: Windows 95/98/2000',
    'Topic :: System :: Systems Administration'
]

base_dir = os.path.dirname(__file__)

DUNDER_ASSIGN_RE = re.compile(r"""^__\w+__\s*=\s*['"].+['"]$""")
about = {}
with open(os.path.join(base_dir, "wmi.py"), "rb") as f:
    for line in f.read().decode("utf-8").splitlines():
        if DUNDER_ASSIGN_RE.search(line):
            exec(line, about)
changes = ""

TO_STRIP = set([":class:", ":mod:", ":meth:", ":func:", ":doc:"])
with open(os.path.join(base_dir, "README.rst"), "rb") as f:
    readme = f.read().decode("utf-8")
    for s in TO_STRIP:
        readme = readme.replace(s, "")


install_requires = [
    "pywin32"
]
extras_require = {
    "tests": [
        "pytest",
    ],
    "docs": ["sphinx"],
    "package": [
        # Wheel building and PyPI uploading
        "wheel",
        "twine",
    ],
}
extras_require["dev"] = (
    extras_require["tests"]
    + extras_require["docs"]
    + extras_require["package"]
)
extras_require["all"] = list(
    {req for extra, reqs in extras_require.items() for req in reqs}
)

setup (
    name=about["__title__"],
    version=about["__version__"],
    description=about["__description__"],
    long_description="{}\n\n{}".format(readme, changes),
    long_description_content_type = "text/x-rst",
    author=about["__author__"],
    author_email=about["__email__"],
    url=about["__url__"],
    license=about["__license__"],
    py_modules = ["wmi"],
    install_requires=install_requires,
    extras_require=extras_require,
    scripts = ["wmitest.py", "wmiweb.py", "wmitest.cmd", "wmitest.master.ini"],
    data_files = ["readme.rst"]
)
