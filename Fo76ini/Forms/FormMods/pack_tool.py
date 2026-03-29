from colorama import Fore, Back, Style, init
import os, shutil
import subprocess
from pathlib import Path
from shutil import which
import winreg
import argparse
init()

PROJECT_GIT_DIR       = Path(__file__).parent.resolve()

TARGET_BASE_DIR       = PROJECT_GIT_DIR / "Publish"
SOLUTION_PATH         = PROJECT_GIT_DIR / "Fo76ini" / "Fo76ini.sln"
PROGRAM_BIN_DIR       = PROJECT_GIT_DIR / "Fo76ini" / "bin"
EXECUTABLE_NAME       = "Fo76ini.exe"
EXECUTABLE_PATH       = PROGRAM_BIN_DIR / "Release" / EXECUTABLE_NAME
UPDATER_BIN_DIR       = PROJECT_GIT_DIR / "Fo76ini_Updater" / "bin"
UPDATER_SOLUTION_PATH = PROJECT_GIT_DIR / "Fo76ini_Updater" / "Fo76ini_Updater.sln"
DEPENDENCIES_DIR      = PROJECT_GIT_DIR / "Additional files"
VERSION_PATH          = PROJECT_GIT_DIR / "VERSION"
SETUP_ISS_PATH        = PROJECT_GIT_DIR / "setup.iss"

VERSION = "x.x.x"

def get_binaries_path():
    return TARGET_BASE_DIR / f"v{VERSION}"

def get_msbuild_path():
    """Attempts to run 'which', then check common paths, and if all else fails, reads the registry and returns the path to MSBuild.exe as string or None."""
    if which("msbuild") is not None:
        return which("msbuild")

    # Use vswhere to find the latest Visual Studio installation with MSBuild
    program_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")
    vswhere_path = os.path.join(program_files_x86, "Microsoft Visual Studio", "Installer", "vswhere.exe")
    if os.path.exists(vswhere_path):
        try:
            # Find the latest installation of VS with MSBuild component
            cmd = [vswhere_path, "-latest", "-products", "*", "-requires", "Microsoft.Component.MSBuild", "-property", "installationPath"]
            vs_path = subprocess.check_output(cmd, text=True, encoding='utf-8').strip()
            if vs_path:
                msbuild_path = os.path.join(vs_path, "MSBuild", "Current", "Bin", "MSBuild.exe")
                if os.path.isfile(msbuild_path):
                    return msbuild_path
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass # Fallback to other methods

    # Fallback to registry check for older versions
    try:
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as reg:
            with winreg.OpenKey(reg, r"SOFTWARE\Microsoft\MSBuild\ToolsVersions\4.0") as key:
                msbuild_path, _ = winreg.QueryValueEx(key, "MSBuildToolsPath")
                path = os.path.join(msbuild_path, "MSBuild.exe")
                if os.path.isfile(path):
                    return path
    except (FileNotFoundError, OSError):
        pass

    return None

def get_7zip_path():
    return which("7z") or which("7za")

def get_version():
    global VERSION
    try:
        with open(VERSION_PATH, "r") as f:
            VERSION = f.read().strip()
    except FileNotFoundError:
        print(Fore.RED + f"ERROR: Version file not found at '{VERSION_PATH}'" + Fore.RESET)
    except IOError as e:
        print(Fore.RED + f"ERROR: Couldn't read VERSION file: {e}" + Fore.RESET)

def set_version():
    global VERSION
    try:
        VERSION = input("VERSION: ")
        with open(VERSION_PATH, "w") as f:
            f.write(VERSION + "\n")
        print("Version set.")
    except KeyboardInterrupt:
        print("\nAbort.")
        return

def restore_nuget():
    nuget_path = which("nuget")
    if nuget_path is None:
        print(Fore.RED + "ERROR: NuGet not found! Please install NuGet and add it to your PATH." + Fore.RESET)
        return

    subprocess.run([nuget_path, "restore", str(SOLUTION_PATH)])
    subprocess.run([nuget_path, "restore", str(UPDATER_SOLUTION_PATH)])

def update_nuget_packages():
    print("Updating NuGet packages to their latest versions...")
    nuget_path = which("nuget")
    if nuget_path is None:
        print(Fore.RED + "ERROR: NuGet not found! Please install NuGet and add it to your PATH." + Fore.RESET)
        return

    # This will update all packages in the solution to their latest stable version.
    print(f"Updating packages for {str(SOLUTION_PATH)}...")
    subprocess.run([nuget_path, "update", str(SOLUTION_PATH)])
    print(f"Updating packages for {str(UPDATER_SOLUTION_PATH)}...")
    subprocess.run([nuget_path, "update", str(UPDATER_SOLUTION_PATH)])
    print("NuGet package update process finished.")

def install_dependencies():
    print("Installing dependencies using Scoop...")
    scoop_path = which("scoop")
    if scoop_path is None:
        print(Fore.RED + "ERROR: Scoop not found! Please install Scoop first (Run 'irm get.scoop.sh | iex' in PowerShell)." + Fore.RESET)
        return

    print("Adding 'extras' bucket (for rcedit)...")
    subprocess.run([scoop_path, "bucket", "add", "extras"])
    print("Installing tools...")
    subprocess.run([scoop_path, "install", "7zip", "git", "rcedit", "inno-setup", "pandoc", "nuget"])

def build_updater(debug = False):
    print("Building updater...")
    configuration = "Debug" if debug else "Release"
    msbuild_path = get_msbuild_path()
    if msbuild_path is None:
        print(Fore.RED + "ERROR: MSBuild not found!" + Fore.RESET)
        return
    subprocess.run([msbuild_path, str(SOLUTION_PATH), f"/p:Configuration={configuration}", "/t:Fo76ini_Updater"])
    if debug:
        copytree(str(UPDATER_BIN_DIR / configuration), str(PROGRAM_BIN_DIR / configuration))
    else:
        copytree(str(UPDATER_BIN_DIR / configuration), str(get_binaries_path()))

def build_app(debug = False):
    print("Building app...")
    configuration = "Debug" if debug else "Release"
    msbuild_path = get_msbuild_path()
    if msbuild_path is None:
        print(Fore.RED + "ERROR: MSBuild not found!" + Fore.RESET)
        return
    subprocess.run([msbuild_path, str(SOLUTION_PATH), f"/p:Configuration={configuration}", "/t:Fo76ini"])
    if not debug:
        copytree(str(PROGRAM_BIN_DIR / "Release"), str(get_binaries_path()))

def copy_additions(debug = False):
    print("Copying additional files...")
    if debug:
        copytree(str(DEPENDENCIES_DIR), str(PROGRAM_BIN_DIR / "Debug"))
    else:
        copytree(str(DEPENDENCIES_DIR), str(get_binaries_path()))

def pack_release():
    print("Packing to v{0}.zip...".format(VERSION))
    sevenzip_path = get_7zip_path()
    if sevenzip_path is None:
        print(Fore.RED + "ERROR: 7-Zip not found!" + Fore.RESET)
        return
    archive_path = str(TARGET_BASE_DIR / f"v{VERSION}.zip")
    source_path = str(get_binaries_path() / "*")
    subprocess.run([sevenzip_path, "a", archive_path, source_path])
    print("Done.")

def use_rcedit():
    print("Setting executable version to '{0}'...".format(VERSION))
    rcedit_path = which("rcedit")
    if rcedit_path is None:
        print(Fore.RED + "ERROR: rcedit not found!" + Fore.RESET)
        return
    subprocess.run([rcedit_path, str(get_binaries_path() / EXECUTABLE_NAME), "--set-file-version", VERSION, "--set-product-version", VERSION])

def update_inno():
    print("Changing version number in setup.iss ...")
    content = ""
    with open(SETUP_ISS_PATH, "r") as f:
        for line in f:
            if line.startswith("#define ProjectVersion"):
                line = "#define ProjectVersion \"" + VERSION + "\"\n"
                print("Line changed: " + line, end="")
            if line.startswith("#define MyAppExeName"):
                line = "#define MyAppExeName \"" + EXECUTABLE_NAME + "\"\n"
                print("Line changed: " + line, end="")
            if line.startswith("#define ProjectGitDir"):
                line = "#define ProjectGitDir \"" + str(PROJECT_GIT_DIR).rstrip("\\") + "\"\n"
                print("Line changed: " + line, end="")
            #if line.startswith("#define ProjectPackTargetDir"):
            #    line = "#define ProjectPackTargetDir \"" + TARGET_BASE_DIR.rstrip("\\") + "\"\n"
            #    print("Line changed: " + line, end="")
            content += line
    with open(SETUP_ISS_PATH, "w") as f:
        f.write(content)

def build_inno():
    print("Building setup using ISCC...")
    iscc_path = which("iscc")
    if iscc_path is None:
        print(Fore.RED + "ERROR: ISCC (Inno Setup) not found!" + Fore.RESET)
        return
    subprocess.run([iscc_path, str(SETUP_ISS_PATH)])

def convert_md():
    print("Converting Markdown to HTML and RTF")
    pandoc_path = which("pandoc")
    if pandoc_path is None:
        print(Fore.RED + "ERROR: Pandoc not found!" + Fore.RESET)
        return
    subprocess.run([pandoc_path, "--standalone", "-f", "gfm", "What's new.md", "-o", "whatsnew.html", "--css=Pandoc/pandoc-style.css", "-H", "Pandoc/pandoc-header.html"])
    subprocess.run([pandoc_path, "--standalone", "-f", "gfm", "What's new.md", "-o", "whatsnewdark.html", "--css=Pandoc/pandoc-style-dark.css", "-H", "Pandoc/pandoc-header.html"])
    subprocess.run([pandoc_path, "--standalone", "What's new.md", "-o", "What's new.rtf"])

def open_dir():
    if os.path.exists(TARGET_BASE_DIR):
        os.startfile(str(TARGET_BASE_DIR))
    else:
        print("ERROR: Path does not exist.")

# https://stackoverflow.com/a/7550424
def mkdir(newdir):
    """Create a directory, including parent directories, if it doesn't exist."""
    os.makedirs(newdir, exist_ok=True)

# https://stackoverflow.com/a/7550424
def copytree(src, dst, symlinks=False):
    """Recursively copy a directory tree using copy2().

    The destination directory must not already exist.
    If exception(s) occur, an Error is raised with a list of reasons.

    If the optional symlinks flag is true, symbolic links in the
    source tree result in symbolic links in the destination tree; if
    it is false, the contents of the files pointed to by symbolic
    links are copied.

    XXX Consider this example code rather than the ultimate tool.

    """
    names = os.listdir(src)
    mkdir(dst)
    errors = []
    for name in names:
        srcname = os.path.join(src, name)
        dstname = os.path.join(dst, name)
        try:
            if symlinks and os.path.islink(srcname):
                linkto = os.readlink(srcname)
                os.symlink(linkto, dstname)
            elif os.path.isdir(srcname):
                copytree(srcname, dstname, symlinks)
            else:
                shutil.copy2(srcname, dstname)
            # XXX What about devices, sockets etc.?
        except (IOError, os.error) as why:
            errors.append((srcname, dstname, str(why)))
        # catch the Error from the recursive copytree so that we can
        # continue with other files
        except Exception as err:
            errors.extend(err.args[0])
    try:
        shutil.copystat(src, dst)
    except OSError:
        # can't copy file access times on Windows, or other errors
        pass

def run_interactive():
    print("""-----------------------------------------
                Pack Tool""")

    warn_text = get_warn_text()
    if warn_text:
        print("-----------------------------------------\n" + warn_text + Fore.RESET, end="")
    else:
        print("-----------------------------------------\n" + Fore.GREEN + "All requirements found!\n" + Fore.RESET, end="")

    while True:
        print("-----------------------------------------")
        print("You can also use command line arguments!\nSee: " + Fore.MAGENTA + "$ " + Fore.BLUE + "python pack_tool.py --help" + Fore.RESET)
        print("-----------------------------------------")
        print(f"""{Fore.BLUE}Set version
{Fore.MAGENTA}(1){Fore.RESET} Set "VERSION" (current: {Fore.GREEN}{VERSION}{Fore.RESET})

{Fore.BLUE}Building
{Fore.MAGENTA}(2){Fore.RESET} Restore NuGet packages
{Fore.MAGENTA}(i){Fore.RESET} Install build dependencies (via Scoop)
{Fore.MAGENTA}(3){Fore.RESET} Update NuGet packages
{Fore.MAGENTA}(4){Fore.RESET} Build app (Debug)
{Fore.MAGENTA}(5){Fore.RESET} Build app (Release)
{Fore.MAGENTA}(6){Fore.RESET} Pack app to *.zip
{Fore.MAGENTA}(7){Fore.RESET} Build setup

{Fore.BLUE}What's new.md
{Fore.MAGENTA}(8){Fore.RESET} Convert Markdown to HTML and RTF using Pandoc

{Fore.BLUE}Others
{Fore.MAGENTA}(9){Fore.RESET} Open target folder
{Fore.MAGENTA}(0){Fore.RESET} Exit (Ctrl+C)
-----------------------------------------""")
        try:
            i = input(">>> ").strip()
        except KeyboardInterrupt:
            print("""^C - Bye bye!
-----------------------------------------""")
            break

        if i == "1":
            set_version()
            # use_rcedit()
        elif i == "2":
            restore_nuget()
        elif i == "i":
            install_dependencies()
        elif i == "3":
            update_nuget_packages()
        elif i == "4":
            build_updater(debug=True)
            build_app(debug=True)
            copy_additions(debug=True)
        elif i == "5":
            build_updater()
            build_app()
            copy_additions()
            use_rcedit()
        elif i == "6":
            pack_release()
        elif i == "7":
            update_inno()
            build_inno()
        elif i == "8":
            convert_md()
        elif i == "9":
            open_dir()
        elif i == "0" or i == "":
            print("""Bye bye!
-----------------------------------------""")
            break
        else:
            print("Input not recognized.")

def run_args(args):
    if args.set_version:
        set_version()
        # use_rcedit()
    if args.install_deps:
        install_dependencies()
    if args.update:
        update_nuget_packages()
    if args.restore:
        restore_nuget()
    if args.build_debug:
        build_updater(debug=True)
        build_app(debug=True)
        copy_additions(debug=True)
    if args.build:
        build_updater()
        build_app()
        copy_additions()
        use_rcedit()
    if args.pack:
        pack_release()
    if args.build_setup:
        update_inno()
        build_inno()
    if args.whatsnew:
        convert_md()

def get_warn_text():
    warn_text = ""
    if not os.path.exists(PROJECT_GIT_DIR):
        warn_text += Fore.YELLOW + f"WARN: Project folder '{PROJECT_GIT_DIR}' doesn't exist!\n"
    if not os.path.isdir(os.path.join(PROJECT_GIT_DIR, "Fo76ini")):
        warn_text += Fore.YELLOW + "WARN: " + Fore.RESET + "\"Fo76ini\" folder doesn't exist!\n      Please run the script within the git repo folder."
    if which("rcedit") is None:
        warn_text += Fore.YELLOW + "WARN: " + Fore.RESET + "rcedit not found!\n"
    if get_7zip_path() is None:
        warn_text += Fore.YELLOW + "WARN: " + Fore.RESET + "7-Zip not found!\n"
    if which("iscc") is None:
        warn_text += Fore.YELLOW + "WARN: " + Fore.RESET + "ISCC (Inno Setup Compiler) not found!\n"
    if which("pandoc") is None:
        warn_text += Fore.YELLOW + "WARN: " + Fore.RESET + "Pandoc not found!\n"
    if get_msbuild_path() is None:
        warn_text += Fore.YELLOW + "WARN: " + Fore.RESET + "MSBuild not found!\n"
    if which("nuget") is None:
        warn_text += Fore.YELLOW + "WARN: " + Fore.RESET + "NuGet not found!\n"

    if warn_text:
        warn_text += Fore.YELLOW + "\nBuilding might fail if requirements are missing.\nMake sure you installed them properly and added them to your PATH.\n\nYou can install most of them with scoop like so:\n> " + Fore.BLUE + "scoop install 7zip git rcedit inno-setup pandoc\n"

    return warn_text

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Helper script for building Fallout 76 Quick Configuration')
    parser.add_argument('-v', '--set-version', help='set the current version', required=False, action='store_true')
    parser.add_argument('-i', '--install-deps', help='install dependencies using Scoop', required=False, action='store_true')
    parser.add_argument('-u', '--update', help='update nuget packages to their latest version', required=False, action='store_true')
    parser.add_argument('-r', '--restore', help='restore nuget packages', required=False, action='store_true')
    parser.add_argument('-b', '--build', help='build the app and updater', required=False, action='store_true')
    parser.add_argument('-d', '--build-debug', help='build the app and updater (Debug configuration)', required=False, action='store_true')
    parser.add_argument('-p', '--pack', help='pack the app into a zip archive', required=False, action='store_true')
    parser.add_argument('-s', '--build-setup', help='build the setup', required=False, action='store_true')
    parser.add_argument('-w', '--whatsnew', help='update the "What\'s new?" files', required=False, action='store_true')
    args = parser.parse_args()

    mkdir(str(TARGET_BASE_DIR))
    get_version()

    args_list = [args.install_deps, args.update, args.restore, args.build_debug, args.build, args.build_setup, args.pack, args.set_version, args.whatsnew]
    #if args_list.count(True) > 1:
    #    print("ERROR: Only one argument allowed")
    if args_list.count(True) >= 1:
        run_args(args)
    else:
        run_interactive()