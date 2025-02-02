import subprocess
import sys

def install_requirements():
    requirements = [
        'pyautogui',
        'pillow',
        'opencv-python',
        'keyboard',
        'pywin32',
        'selenium',
        'webdriver_manager',
        'numpy',
        'tkinter',
        'json',
        'requests'
    ]
    
    print("Installing MacroFlow Studio requirements...")
    
    for package in requirements:
        print(f"Installing {package}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"Successfully installed {package}")
        except subprocess.CalledProcessError:
            print(f"Failed to install {package}")
    
    print("\nInstallation complete!")

if __name__ == "__main__":
    install_requirements()