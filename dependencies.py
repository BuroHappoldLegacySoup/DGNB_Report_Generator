from importlib.metadata import distribution, PackageNotFoundError
import subprocess

REQUIRED_PACKAGES = [
    'python_calamine',
    'python-docx',
]

def check_packages():
    for package in REQUIRED_PACKAGES:
        try:
            dist = distribution(package)
            #print(f'{package} ({dist.version}) is installed')
        except PackageNotFoundError:
            print(f'{package} is missing!')
            install = input('Would you like to install the missing packages (y/n)?')
            if install.lower() == 'y':
                subprocess.call(['pip', 'install', package])
            else:
                print('Sorry. The program won\'t work without the necessary packages.')
