import os
import sys
import shutil
import subprocess
import glob
import logging
import platform
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def clean_build():
    """Clean previous build files"""
    try:
        dirs_to_clean = ['build', 'dist']
        files_to_clean = ['*.spec']
        
        for dir_name in dirs_to_clean:
            if os.path.exists(dir_name):
                logger.info(f"Removing directory: {dir_name}")
                try:
                    shutil.rmtree(dir_name)
                except Exception as e:
                    logger.warning(f"Failed to remove {dir_name}: {str(e)}")
                    # Try to remove individual files if directory removal fails
                    for root, dirs, files in os.walk(dir_name, topdown=False):
                        for name in files:
                            try:
                                os.remove(os.path.join(root, name))
                            except Exception as e:
                                logger.warning(f"Failed to remove file {name}: {str(e)}")
                        for name in dirs:
                            try:
                                os.rmdir(os.path.join(root, name))
                            except Exception as e:
                                logger.warning(f"Failed to remove directory {name}: {str(e)}")
                    try:
                        os.rmdir(dir_name)
                    except Exception as e:
                        logger.warning(f"Failed to remove empty directory {dir_name}: {str(e)}")
        
        for pattern in files_to_clean:
            for file in glob.glob(pattern):
                logger.info(f"Removing file: {file}")
                try:
                    os.remove(file)
                except Exception as e:
                    logger.warning(f"Failed to remove file {file}: {str(e)}")
                
        logger.info("Clean build completed successfully")
    except Exception as e:
        logger.error(f"Error during clean build: {str(e)}")
        raise

def build_app():
    """Build the app for the current platform"""
    try:
        # Detect platform
        system = platform.system()
        arch = platform.machine()
        
        if system == "Windows":
            build_windows_exe()
        elif system == "Darwin":  # macOS
            build_macos(arch)
        else:
            logger.error(f"Unsupported platform: {system}")
            raise Exception(f"Building for {system} is not supported yet")
            
    except Exception as e:
        logger.error(f"Build process failed: {str(e)}")
        raise

def build_windows_exe():
    """Build the Windows executable"""
    try:
        # Clean previous builds
        clean_build()
        
        # Create version info
        version = datetime.now().strftime("%Y.%m.%d")
        logger.info(f"Building Windows executable version: {version}")
        
        # Check if PyInstaller is installed
        try:
            import PyInstaller
            logger.info(f"PyInstaller version: {PyInstaller.__version__}")
        except ImportError:
            logger.error("PyInstaller not found. Installing...")
            subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
        
        # Create spec file first with Windows-specific settings
        spec_cmd = [
            sys.executable,
            "-m",
            "PyInstaller",
            '--name=CSV_Search_App',
            '--windowed',  # No console window
            '--clean',
            '--noconfirm',
            '--onefile',  # Create a single executable
            '--win-private-assemblies',  # Include private assemblies
            f'--version-file=version_info.txt',
            '--add-data=requirements.txt;.',  # Windows uses semicolon
            '--specpath=.',
            'Code_V1.py'
        ]
        
        logger.info("Creating spec file...")
        result = subprocess.run(spec_cmd, capture_output=True, text=True)
        
        if result.returncode != 0:
            logger.error(f"Spec file creation failed:\n{result.stderr}")
            raise Exception("Spec file creation failed")
        
        # Build using the spec file
        build_cmd = [
            sys.executable,
            "-m",
            "PyInstaller",
            '--clean',
            '--noconfirm',
            'CSV_Search_App.spec'
        ]
        
        logger.info("Starting PyInstaller build...")
        result = subprocess.run(build_cmd, capture_output=True, text=True)
        
        if result.returncode != 0:
            logger.error(f"PyInstaller build failed:\n{result.stderr}")
            raise Exception("PyInstaller build failed")
        
        # Create release directory
        release_dir = f'release_windows_{version}'
        if not os.path.exists(release_dir):
            os.makedirs(release_dir)
            logger.info(f"Created release directory: {release_dir}")
        
        # Copy files to release directory
        files_to_copy = [
            ('dist/CSV_Search_App.exe', 'CSV_Search_App.exe'),
            ('README.md', 'README.md'),
            ('LICENSE', 'LICENSE'),
            ('requirements.txt', 'requirements.txt')
        ]
        
        for src, dst in files_to_copy:
            if os.path.exists(src):
                shutil.copy(src, os.path.join(release_dir, dst))
                logger.info(f"Copied {src} to release directory")
            else:
                logger.warning(f"Source file not found: {src}")
        
        logger.info(f"\nWindows executable build completed successfully!")
        logger.info(f"Release files are in the '{release_dir}' directory")
        
    except Exception as e:
        logger.error(f"Windows executable build failed: {str(e)}")
        raise

def build_macos(arch):
    """Build the macOS application"""
    try:
        # Clean previous builds
        clean_build()
        
        # Create version info
        version = datetime.now().strftime("%Y.%m.%d")
        logger.info(f"Building macOS version: {version}")
        
        # Check if PyInstaller is installed
        try:
            import PyInstaller
            logger.info(f"PyInstaller version: {PyInstaller.__version__}")
        except ImportError:
            logger.error("PyInstaller not found. Installing...")
            subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
        
        logger.info(f"Building for architecture: {arch}")
        
        # macOS uses colon as path separator for PyInstaller
        path_sep = ":"
        
        # Create a custom .spec file instead of using command-line args
        spec_content = f"""
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['Code_V1.py'],
    pathex=[],
    binaries=[],
    datas=[('requirements.txt', '.')],
    hiddenimports=[
        'PIL._tkinter_finder',
        'pandas',
        'polars',
        'PyQt6',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
        'PyQt6.QtWidgets',
        'PyQt6.sip',
        'matplotlib',
        'matplotlib.backends.backend_qt5agg',
        'matplotlib.backends.backend_qt5',
        'matplotlib.backends.backend_agg',
        'matplotlib.backends.backend_svg',
        'matplotlib.backends.backend_pdf',
        'matplotlib.backends.backend_ps',
        'seaborn',
        'numpy',
        'openpyxl',
        'xlsxwriter',
        'reportlab',
        'psutil'
    ],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='CSV_Search_App',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CSV_Search_App',
)

app = BUNDLE(
    coll,
    name='CSV_Search_App.app',
    icon=None,
    bundle_identifier=None,
)
"""
        
        # Write the spec file
        with open('CSV_Search_App.spec', 'w') as f:
            f.write(spec_content)
        
        logger.info("Created custom spec file")
        
        # Build using the spec file
        build_cmd = [
            sys.executable,
            "-m",
            "PyInstaller",
            '--clean',
            '--noconfirm',
            'CSV_Search_App.spec'
        ]
        
        logger.info("Starting PyInstaller build...")
        # Run the command and capture real-time output
        process = subprocess.Popen(
            build_cmd, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1
        )
        
        # Print output in real-time
        for line in iter(process.stdout.readline, ''):
            logger.info(line.strip())
        
        process.stdout.close()
        return_code = process.wait()
        
        if return_code != 0:
            logger.error(f"PyInstaller build failed with return code: {return_code}")
            raise Exception("PyInstaller build failed")
        
        # Create release directory
        release_dir = f'release_macos_{arch}_{version}'
        if not os.path.exists(release_dir):
            os.makedirs(release_dir)
            logger.info(f"Created release directory: {release_dir}")
        
        # For macOS, we need to copy the entire .app bundle
        app_path = 'dist/CSV_Search_App.app'
        if os.path.exists(app_path):
            target_path = os.path.join(release_dir, 'CSV_Search_App.app')
            # Use shutil.copytree for directories
            if os.path.exists(target_path):
                shutil.rmtree(target_path)
            shutil.copytree(app_path, target_path)
            logger.info(f"Copied {app_path} to release directory")
            
            # Copy additional files
            extra_files = [
                ('README.md', 'README.md'),
                ('LICENSE', 'LICENSE'),
                ('requirements.txt', 'requirements.txt')
            ]
            
            for src, dst in extra_files:
                if os.path.exists(src):
                    shutil.copy(src, os.path.join(release_dir, dst))
                    logger.info(f"Copied {src} to release directory")
                else:
                    logger.warning(f"Source file not found: {src}")
        else:
            logger.error(f"App bundle not found: {app_path}")
            raise Exception("App bundle not found")
        
        logger.info(f"\nmacOS application build completed successfully!")
        logger.info(f"Release files are in the '{release_dir}' directory")
        
    except Exception as e:
        logger.error(f"macOS build failed: {str(e)}")
        raise

if __name__ == '__main__':
    build_app()