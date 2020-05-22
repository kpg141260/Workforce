import fnmatch
from setuptools import find_packages, setup
from setuptools.command.build_py import build_py as build_py_orig

excluded = [
    '.git*',
    '.vscode',
    '*workspace',
]

class build_py(build_py_orig):
    def find_package_modules(self, package, package_dir):
        modules = super().find_package_modules(package, package_dir)
        return [
            (pkg, mod, file)
            for (pkg, mod, file) in modules
            if not any(fnmatch.fnmatchcase(file, pat=pattern) for pattern in excluded)
        ]

setup(name='Workforce_PY',
    classifiers=[
        'Development Status :: 4 - Beta',
        'Environment :: Web Environment',
        'Intended Audience :: Developers',
        'Operating System :: OS Independent',
        'License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)',
        'Programming Language :: Python :: 3.8',
        'Topic :: Office/Business :: Scheduling',
        'Topic :: Software Development :: Libraries',
        ],
    python_requires='>=3.8',
    version='0.1',
    author='Peter Gossler',
    author_email='kpg141260@live.de',
    description='A Library for Workforce planning in contact centers.',
    long_description_content_type="text/markdown",
    url='https://github.com/kpg141260/erlang',
    packages=find_packages() , 
    long_description=open('README.md').read(),
    cmdclass={'build_py': build_py},
    zip_safe=True
)