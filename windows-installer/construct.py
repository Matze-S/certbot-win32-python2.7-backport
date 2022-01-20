#!/usr/bin/env python3
import contextlib
import ctypes
import os
import shutil
import struct
import subprocess
import sys
import tempfile
import time
try:
    from pathlib import Path
except ImportError:
    from pathlib2 import Path  # Backport
    
PYTHON_VERSION = (2, 7, 18)
PYTHON_BITNESS = 32
PYWIN32_VERSION = 227  # do not forget to edit pywin32 dependency accordingly in setup.py
NSIS_VERSION = '3.04'

def get_cache_dir(ensure_existence=False):
    if os.name == 'posix' and sys.platform != 'darwin':
        # Linux, Unix, AIX, etc.
        # use ~/.cache if empty OR not set
        xdg = os.environ.get("XDG_CACHE_HOME", None) or (os.path.expanduser('~/.cache'))
        p = Path(xdg, 'pynsist')

    elif sys.platform == 'darwin':
        p = Path(os.path.expanduser('~'), 'Library/Caches/pynsist')

    else:
        # Windows (hopefully)
        local = os.environ.get('LOCALAPPDATA', None) or (os.path.expanduser('~\\AppData\\Local'))
        if local.startswith('~'):
            logger.warning("Could not find cache directory. Please set any of "
                           "these environment variables: "
                           "LOCALAPPDATA, HOME, USERPROFILE or HOMEPATH")
        p = Path(local, 'pynsist')

    if ensure_existence:
        try:
            p.mkdir(parents=True)
        except OSError as e:
            # Py2 compatible equivalent of FileExistsError
            if e.errno != errno.EEXIST:
                raise

    return p


def main():
    build_path, repo_path, venv_path, venv_python = _prepare_environment()

    _copy_assets(build_path, repo_path)
    
    _prepare_build_tools(venv_path, venv_python, repo_path, build_path)
    _compile_wheels(repo_path, build_path, venv_python)

    installer_cfg_path = _generate_pynsist_config(repo_path, build_path, venv_python)
    nsis_path = os.path.join(build_path, 'nsis')
    os.makedirs(nsis_path)
    for name in os.listdir(os.path.join(repo_path, 'windows-installer')):
        if name.endswith('.msi'):
            shutil.copy(os.path.join(repo_path, 'windows-installer', name), nsis_path)
    shutil.copy(os.path.join(repo_path, 'windows-installer', 'certbot.ico'), build_path)
    _build_installer(installer_cfg_path, venv_path)

    print('Done')


def _build_installer(installer_cfg_path, venv_path):
    print('Build the installer')
    subprocess.check_call([os.path.join(os.path.dirname(sys.executable), 'pynsist'), installer_cfg_path])
#    subprocess.check_call([os.path.join(venv_path, 'Scripts', 'pynsist.exe'), installer_cfg_path])

def _compile_wheels(repo_path, build_path, venv_python):
    print('Compile wheels')

    wheels_path = os.path.join(build_path, 'wheels')
    os.makedirs(wheels_path)
 
    certbot_packages = ['acme', 'certbot']
    # Uncomment following line to include all DNS plugins in the installer
    #certbot_packages.extend([name for name in os.listdir(repo_path) if name.startswith('certbot-dns-')])
    wheels_project = [os.path.join(repo_path, package) for package in certbot_packages]

#    print('Prepare Constraints')
    with _prepare_constraints(repo_path) as constraints_file_path:
#        command = [venv_python, '-m', 'pip', 'wheel', '-w', wheels_path, '--constraint', constraints_file_path]
        command = [venv_python, '-m', 'pip', 'wheel', '-w', wheels_path, '--constraint', 'constraints.txt']
        command.extend(wheels_project)
        subprocess.check_call(command)
        
    print('Copy Wheels to pypi cache') 
    cache_dir = os.path.join(str(get_cache_dir()), 'pypi')
    for filename in os.listdir(wheels_path):
      if filename.endswith('.whl'):
        version = filename.split('-', 2)[1]
        name    = filename.split('-', 1)[0]
        cache_subdir = os.path.join(cache_dir, name)
        cache_subdir = os.path.join(cache_subdir, version)
        if not os.path.exists(cache_subdir):
          os.makedirs(cache_subdir);
        shutil.copy(os.path.join(wheels_path, filename), cache_subdir)
        print os.path.join(cache_subdir, filename)
    
    print('Install Wheels')
    wheels_files = [os.path.join(wheels_path, name) for name in os.listdir(wheels_path) if name.endswith('.whl')]
    command = [venv_python, '-m', 'pip', 'install']
    command.extend(wheels_files)
    subprocess.check_call(command);

def _prepare_build_tools(venv_path, venv_python, repo_path, build_path):
    print('Prepare build tools')
    #subprocess.check_call([sys.executable, '-m', 'venv', venv_path])
    subprocess.check_call(['virtualenv', venv_path])

    ######################
    #wheels_path = os.path.join(build_path, 'wheels')
    #os.makedirs(wheels_path)
    #
    #own_packages = ['../pynsist']
    #wheels_project = [os.path.join(repo_path, package) for package in own_packages]
    #
    #with _prepare_constraints(repo_path) as constraints_file_path:
    #    command = [venv_python, '-m', 'pip', 'wheel', '-w', wheels_path, '--constraint', constraints_file_path]
    #    command.extend(wheels_project)
    #    subprocess.check_call(command)
    ######################
    
    subprocess.check_call([venv_python, os.path.join(repo_path, 'letsencrypt-auto-source', 'pieces', 'pipstrap.py')])
    #subprocess.check_call([venv_python, os.path.join(repo_path, 'tools', 'pip_install.py'), 'pynsist' ])

    subprocess.check_call([venv_python, '-m', 'ensurepip', '--upgrade'])
#    subprocess.check_call([venv_python, '-m', 'pip', 'install', 'pynsist<2'])
#    subprocess.check_call([venv_python, '-m', 'pip', 'install', os.path.join(repo_path, "windows-installer/./pynsist-1.12-py2.py3-none-any.whl") ])
#    subprocess.check_call([venv_python, '-m', 'pip', 'install', os.path.join(repo_path, "windows-installer/./pynsist-2.5.1-py2.py3-none-any.whl") ])
#    subprocess.check_call([venv_python, '-m', 'pip', 'install', 'pynsist', os.path.join(repo_path, '../pynsist') ])
    #subprocess.check_call([venv_python, os.path.join(repo_path, 'tools', 'pip_install.py'), 'pynsist', os.path.join(repo_path, '../pynsist') ])

    #subprocess.check_call(['choco', 'upgrade', '--allow-downgrade', '-y', 'nsis', '--version', NSIS_VERSION])



@contextlib.contextmanager
def _prepare_constraints(repo_path):
#    requirements = os.path.join(repo_path, 'letsencrypt-auto-source', 'pieces', 'dependency-requirements.txt')
#    constraints = subprocess.check_output(
#        [sys.executable, os.path.join(repo_path, 'tools', 'strip_hashes.py'), requirements],
#        universal_newlines=True)
#    constraints = subprocess.check_output(
#        [sys.executable, '-m', 'pip', 'freeze'],
#        universal_newlines=True)
    constraints = subprocess.check_output(
        ['cat', 'constraints.txt'],
        universal_newlines=True)
    workdir = tempfile.mkdtemp()
    try:
        constraints_file_path = os.path.join(workdir, 'constraints.txt')
        with open(constraints_file_path, 'a') as file_h:
            file_h.write(constraints)
            file_h.write('pywin32=={0}'.format(PYWIN32_VERSION))
        yield constraints_file_path
    finally:
        shutil.rmtree(workdir)


def _copy_assets(build_path, repo_path):
    print('Copy assets')
    if os.path.exists(build_path):
        os.rename(build_path, '{0}.{1}.bak'.format(build_path, int(time.time())))
    os.makedirs(build_path)
    shutil.copy(os.path.join(repo_path, 'windows-installer', 'certbot.ico'), build_path)
    shutil.copy(os.path.join(repo_path, 'windows-installer', 'run.bat'), build_path)
    shutil.copy(os.path.join(repo_path, 'windows-installer', 'template.nsi'), build_path)
    shutil.copy(os.path.join(repo_path, 'windows-installer', 'renew-up.ps1'), build_path)
    shutil.copy(os.path.join(repo_path, 'windows-installer', 'renew-down.ps1'), build_path)


def _generate_pynsist_config(repo_path, build_path, venv_python):
    print('Generate pynsist configuration')

    wheels_path = os.path.join(build_path, 'wheels');
    wheels_packages = [name.split('-', 1)[0] for name in os.listdir(wheels_path) if name.endswith('.whl')]
    wheels_packages = [name.split('.', 1)[0] for name in wheels_packages]

    wheels_files = [name for name in os.listdir(wheels_path) if name.endswith('.whl')]
    wheels_packages2 = [name.split('-', 1)[0] for name in wheels_files]
    wheels_packages2 = [name.split('.', 1)[0] for name in wheels_packages2]
    wheels_packages2 = [name.replace('_', '-') for name in wheels_packages2]

    pypi_wheels = ['=='.join(name.split('-', 2)[0:2]) for name in os.listdir(wheels_path) if name.endswith('.whl') and not name.startswith('certbot') and not name.startswith('acme') and not name.startswith('letsencryapt')]
    
    import pkg_resources
    # installed_packages = [(pkg.project_name, pkg.version) for pkg in pkg_resources.working_set]
    installed_packages = [pkg.project_name for pkg in pkg_resources.working_set]

    pypkgcmdline = 'import pkg_resources; pkgs = [pkg.project_name for pkg in pkg_resources.working_set]; print \'\\n\'.join(pkgs)'
    packagesoutput = subprocess.check_output([sys.executable, '-c', pypkgcmdline ], universal_newlines=True)
    installed_packages = packagesoutput.split('\n')
   
    print "installed_packages="
    print installed_packages
    
    certbot_packages = ['acme', 'certbot']
    # Uncomment following line to include all DNS plugins in the installer
    #certbot_packages.extend([name.replace('-', '_') for name in os.listdir(repo_path) if name.startswith('certbot-dns-')])
    
    pywin32_paths_file = os.path.join(build_path, 'pywin32_paths.py')

    # Pywin32 uses non-standard folders to hold its packages. We need to instruct pynsist bootstrap
    # explicitly to add them into sys.path. This is done with a custom "pywin32_paths.py" that is
    # referred in the pynsist configuration as an "extra_preamble".
    # Reference example: https://github.com/takluyver/pynsist/tree/master/examples/pywebview
    with open(pywin32_paths_file, 'w') as file_h:
        file_h.write('''\
pkgdir = os.path.join(installdir, 'pkgs')

sys.path.extend([
    os.path.join(pkgdir, 'win32'),
    os.path.join(pkgdir, 'win32', 'lib'),
])

from site import addsitedir
addsitedir(pkgdir)

# Preload pywintypes and pythoncom
pwt = os.path.join(pkgdir, 'pywin32_system32', 'pywintypes{0}{1}.dll')
pcom = os.path.join(pkgdir, 'pywin32_system32', 'pythoncom{0}{1}.dll')
import warnings
with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    import imp
imp.load_dynamic('pywintypes', pwt)
imp.load_dynamic('pythoncom', pcom)
'''.format(PYTHON_VERSION[0], PYTHON_VERSION[1]))

    installer_cfg_path = os.path.join(build_path, 'installer.cfg')

    certbot_pkg_path = os.path.join(repo_path, 'certbot')
    certbot_version = subprocess.check_output([sys.executable, '-c', 'import certbot; print(certbot.__version__)'],
                                              universal_newlines=True, cwd=certbot_pkg_path).strip()

    # If we change the installer name from `certbot-installer-win32.exe`, it should
    # also be changed in tools/create_github_release.py
    with open(installer_cfg_path, 'w') as file_h:
        file_h.write('''\
[Application]
name=Certbot
version={certbot_version}
icon=certbot.ico
publisher=Electronic Frontier Foundation
target=$INSTDIR\\run.bat

[Build]
directory=nsis
;nsi_template=template.nsi
installer_name=certbot-{certbot_version}-installer-{installer_suffix}.exe

[Python]
version={python_version}
bitness={python_bitness}

[Include]
;local_wheels=wheels\\*.whl
pypi_wheels={pypi_wheels}
packages={packages}
files=run.bat
      renew-up.ps1
      renew-down.ps1

[Command certbot]
entry_point=certbot.main:main
extra_preamble=pywin32_paths.py
'''.format(certbot_version=certbot_version,
           installer_suffix='win_amd64' if PYTHON_BITNESS == 64 else 'win32',
           packages='\n         '.join(str(item) for item in certbot_packages),
           pypi_wheels='\n            '.join(str(item) for item in pypi_wheels),
           python_bitness=PYTHON_BITNESS,
           python_version='.'.join(str(item) for item in PYTHON_VERSION)))

        return installer_cfg_path


def _prepare_environment():
    print('Prepare environment')
    #try:
    #    subprocess.check_output(['choco', '--version'])
    #except subprocess.CalledProcessError:
    #    raise RuntimeError('Error: Chocolatey (https://chocolatey.org/) needs '
    #                       'to be installed to run this script.')
    script_path = os.path.realpath(__file__)
    repo_path = os.path.dirname(os.path.dirname(script_path))
    build_path = os.path.join(repo_path, 'windows-installer', 'build')
    venv_path = os.path.join(build_path, 'venv-config')
    venv_python = os.path.join(venv_path, 'Scripts', 'python.exe')

    return build_path, repo_path, venv_path, venv_python


if __name__ == '__main__':
    if os.name != 'nt':
        raise RuntimeError('This script must be run under Windows.')

    if ctypes.windll.shell32.IsUserAnAdmin() == 0:
        # Administrator privileges are required to properly install NSIS through Chocolatey
        raise RuntimeError('This script must be run with administrator privileges.')

    if sys.version_info[:2] != PYTHON_VERSION[:2]:
        raise RuntimeError('This script must be run with Python {0}'
                           .format('.'.join(str(item) for item in PYTHON_VERSION[0:2])))

    if struct.calcsize('P') * 8 != PYTHON_BITNESS:
        raise RuntimeError('This script must be run with a {0} bit version of Python.'
                           .format(PYTHON_BITNESS))
    main()
