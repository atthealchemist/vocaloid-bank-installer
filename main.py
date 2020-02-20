import logging
import os
import platform
import re
import shutil
import subprocess
import winreg
from argparse import ArgumentParser
from pathlib import Path
from tkinter import filedialog, messagebox


class ConsoleLogger:

    def set_level(self, level):
        self.logger.setLevel(level)

    def info(self, message):
        self.logger.info(message)

    def warning(self, message):
        self.logger.warning(message)

    def error(self, message):
        self.logger.error(message)

    def exception(self, message):
        self.logger.exception(message)

    def debug(self, message):
        self.logger.debug(message)

    def log(self, message, level=logging.INFO):
        if level == logging.INFO:
            self.info(message)
        if level == logging.DEBUG:
            self.debug(message)
        if level == logging.ERROR:
            self.error(message)
        if level == logging.WARNING:
            self.warning(message)

    def __init__(self, name=__name__, default_level=logging.DEBUG):
        self.logger = logging.Logger(name)
        if not self.logger.handlers or len(self.logger.handlers) < 1:
            handler = logging.StreamHandler()
            handler.setFormatter(logging.Formatter("[%(name)s]\t %(asctime)s [%(levelname)s]\t %(message)s"))
            handler.setLevel(default_level)
            self.logger.addHandler(handler)


logger = ConsoleLogger("VOCALOID_CRACK")


def is_valid_bank_dir(name):
    return re.match(r"^B+\w+", name)


def get_banks_ids(bank_path):
    banks_ids = []
    for root, dirs, files in os.walk(bank_path):
        for dir_name in filter(lambda d: is_valid_bank_dir(d), dirs):
            banks_ids.append(dir_name)
    return banks_ids


def get_bank_name_by_id(bank_id, source_path):
    result = ""
    logger.debug("Checking source for new banks @ {}".format(source_path))
    for root, dirs, _ in os.walk(source_path):
        for dir in filter(lambda d: d == bank_id, dirs):
            files = os.listdir(r"{}/{}".format(root, dir))
            for file in filter(lambda f: f.endswith("ddb"), files):
                logger.debug("Found ddb file: {}".format(file))
                bank_path = r"{}/{}".format(root, dir)
                bank_name, extension = file.split('.')
                result = bank_name, bank_path
    return result


def move_up_and_cleanup(path):
    source_path = Path(path)
    parent = source_path.parent
    files = os.listdir(path)
    # move level up
    for file in filter(lambda f: os.path.isdir(f), files):
        logger.debug("Moving {} level up ({})".format(file, str(parent)))
        shutil.move(file, str(parent))
    # cleanup old stuff
    if source_path.exists() and not os.listdir(source_path):
        source_path.rmdir()


def is_bank_already_installed(path):
    result = False
    if os.path.exists(path):
        for file in os.listdir(path):
            if str(file).endswith(('.ddb', '.ddi', '.vvd')):
                result = True
    return result


def run_installer(path_to_installer_executable, source_directory):
    installation_result = False
    *_, installer_exe_name = path_to_installer_executable.split('/')
    logger.debug("Got installer executable: {}".format(installer_exe_name))
    installer_name, ext = installer_exe_name.split('.')
    logger.debug("Got installer name: {}".format(installer_name))

    install_dir = r"{}/{}".format(source_directory, installer_name)
    if is_bank_already_installed(install_dir):
        logger.debug("Bank is installed, skipping...")
        installation_result = True
    else:
        logger.debug("Running installer {}, writing @ {}".format(path_to_installer_executable, install_dir))

        installer = subprocess.run([
            path_to_installer_executable,
            '/SP-',
            '/VERYSILENT',
            '/NOCANCEL',
            '/NORESTART',
            '/SUPRESSMSGBOXES',
            '/DIR={}'.format(install_dir)
        ])
        if installer.returncode == 0:
            logger.debug("Successfully installed {} into {}".format(installer_name, source_directory))
            installation_result = True

    move_up_and_cleanup(r"{}/{}/{}".format(source_directory, installer_name, installer_name))

    return installation_result


def get_system_bits():
    bit_dict = dict(
        i386=winreg.KEY_WOW64_32KEY,
        AMD64=winreg.KEY_WOW64_64KEY
    )
    machine = platform.machine()
    return bit_dict.get(machine, winreg.KEY_WOW64_32KEY)


def write_registry(bank_id, bank_name, bank_path):
    logger.debug("Writing registry for {} ({}) @ {}".format(bank_name, bank_id, bank_path))
    bank_path_parent = str(Path(bank_path).parent)
    dict_bank = dict(
        BankName=bank_name.replace(' ', '_'),
        DRP="000055",
        DefaultStyleID="gf22245e-19b1-40e1-a37a-699eaa5b4a1d",
        Key="2ef256c28b85f2f16329acb0a9ba2ea4",
        Name=bank_name,
        Path=bank_path_parent,
        Date="BMRDG7KZR3ZB27DE"
    )

    dict_version = dict(
        Major=4,
        Minor=0,
        Revision=0
    )

    # registry = winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE)

    # install config section
    config_reg_path = r"SOFTWARE\VOCALOID5\Voice\Components\{}".format(bank_id)
    logger.debug("Writing {}".format(config_reg_path))
    try:
        config_key = winreg.OpenKeyEx(
            winreg.HKEY_LOCAL_MACHINE,
            config_reg_path,
            0,
            winreg.KEY_WOW64_64KEY | winreg.KEY_ALL_ACCESS
        )
    except Exception as ex:
        config_key = winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, config_reg_path, 0,
                                        winreg.KEY_WOW64_64KEY | winreg.KEY_ALL_ACCESS)
    # set one key
    for key, value in dict_bank.items():
        logger.debug("Setup key | {} : {}".format(key, value))
        winreg.SetValueEx(config_key, key, 0, winreg.REG_SZ, value)
    winreg.CloseKey(config_key)

    # install version section
    version_reg_path = r'SOFTWARE\VOCALOID5\Voice\Components\{}\Version'.format(bank_id)
    logger.debug("Writing {}".format(version_reg_path))
    winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, version_reg_path)
    try:
        version_key = winreg.OpenKeyEx(
            winreg.HKEY_LOCAL_MACHINE,
            version_reg_path,
            0,
            winreg.KEY_WOW64_64KEY | winreg.KEY_ALL_ACCESS
        )
    except Exception as ex:
        version_key = winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, version_reg_path, 0,
                                         winreg.KEY_WOW64_64KEY | winreg.KEY_ALL_ACCESS)
    for key, value in dict_version.items():
        logger.debug("Setup key | {} : {}".format(key, value))
        winreg.SetValueEx(version_key, key, 0, winreg.REG_DWORD, value)
    winreg.CloseKey(version_key)

    logger.debug("Successfully writing all changes into registry")


def parse_args():
    parser = ArgumentParser()
    parser.add_argument('-s', '--source',
                        action='store',
                        default=r"C:\Program Files\Common Files\VOCALOID5\Voicelib",
                        help="Source directory (where your bank should be installed)")

    parser.add_argument('-i', '--installer',
                        action='store',
                        help="Path to installer (where setup.exe is situated)")

    return parser.parse_args()


def process_bank(source_path, installer_path):
    banks_before_install = set(get_banks_ids(source_path))

    run_installer(installer_path, source_path)

    banks_after_install = set(get_banks_ids(source_path))

    new_banks_ids = list(banks_after_install.difference(banks_before_install))

    logger.debug("New banks ids: {}".format(new_banks_ids))
    for idx, bank_id in enumerate(new_banks_ids):
        logger.debug("Processing bank {}/{} {}".format(idx + 1, len(new_banks_ids), bank_id))
        bank_name, bank_path = get_bank_name_by_id(bank_id, source_path)
        if bank_name:
            write_registry(bank_id, bank_name, bank_path)
            logger.debug("Bank was successfully installed")
            messagebox.showinfo(title="Success", message="Bank {} was successfully installed!".format(bank_name))


def main():
    # elevate(show_console=False)

    installer_path = r"{}".format(filedialog.askopenfilename(title="Select bank installer path (setup.exe)"))
    logger.debug("Set up installer executable @ {}".format(installer_path))

    source_path = r"{}".format(filedialog.askdirectory(title="Select path where you want to place your bank"))
    logger.debug("Set up source path @ {}".format(source_path))

    try:
        process_bank(source_path, installer_path)
    except Exception as ex:
        logger.exception("Exception: {}".format(ex))
        messagebox.showerror(title="Error!", message="Exception: {}".format(ex))


if __name__ == "__main__":
    main()
