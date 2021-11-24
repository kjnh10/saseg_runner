"""Module to contain the EGRunner class and the associated function (which will likely be
deprecated in the future)
"""
# %%
import datetime
import os
import re
import shutil
import subprocess
import sys
import time

from pathlib import Path
from typing import Union

import click
import pythoncom
import win32com.client

SCRIPTDIR_PATH = Path(__file__).parent.resolve()
DEFAULT_PROFILE_NAME = "SAS Asia"
DEFAULT_EG_VERSION = "7.1"


class SASEGRunner:
    """Class that can be used to save settings for running multiple EG files."""

    def __init__(
        self,
        profile_name: str,
        eg_version: str = DEFAULT_EG_VERSION,
        overwrite: bool = False,
        remove_log: bool = True,
        verbose: bool = False,
    ):
        """Class that can be used to save settings for running multiple EG files.

        Args:
            profile_name (str): profile name to use when running EGP files.
            eg_version (str, optional): Which version of EG to use. Defaults to "7.1"
            overwrite (bool, optional): controls whether to save the egp file after exection.
                if False, timestamp is added to filename. Defaults to False
            remove_log (bool, optional): Whether to remove log files or not. Defaults to True
            verbose (bool, optional): [description]. Defaults to False
        """
        self.profile_name = profile_name
        self.eg_version = eg_version
        self.overwrite = overwrite
        self.remove_log = remove_log
        self.verbose = verbose

    def run_egp(
        self,
        egp_path: Union[str, Path],
    ) -> None:
        """Function to run an EGP file from Python.

        Args:
            egp_path (Union[str, Path]): SAS Enterprise Guide file path.
        """
        run_egp(
            egp_path,
            self.profile_name,
            self.eg_version,
            self.overwrite,
            self.remove_log,
            self.verbose,
        )


def run_egp(
    egp_path: Union[str, Path],
    profile_name: str = DEFAULT_PROFILE_NAME,
    eg_version: str = DEFAULT_EG_VERSION,
    overwrite: bool = False,
    remove_log: bool = True,
    verbose: bool = False,
) -> None:
    """Function to run an EGP file from Python.

    Args:
        egp_path (Union[str, Path]): SAS Enterprise Guide file path.
        profile_name (str, optional): profile name to use. Defaults to 'SAS Asia'.
        eg_version (str, optional): Which version of EG to use. Defaults to '7.1'.
        overwrite (bool, optional): controls whether to save the egp file after exection.
            if False, timestamp is added to filename. Defaults to False.
        remove_log (bool, optional): Whether to remove log files or not. Defaults to True.
        verbose (bool, optional): [description]. Defaults to False.
    """
    start_time = time.time()

    egp_path = _check_egp_file_existence(egp_path)
    app = _open_enterprise_guide(eg_version)
    _activate_enterprise_guide_profile(profile_name, app)
    project_object = _open_egp(egp_path, app)
    _run_egp(egp_path, project_object)
    output = _save_and_close_egp(egp_path, overwrite, project_object)
    log_dir = _retrieve_logs(eg_version, verbose, output)
    error_happend = _check_log_for_errors(egp_path, log_dir)
    if remove_log:
        shutil.rmtree(log_dir)
    _finish_and_clean_up(egp_path, output, error_happend)
    elapsed_time = int(time.time() - start_time)
    print(f"elapsed_time:{elapsed_time}[sec]")


def _check_egp_file_existence(egp_path: str) -> Path:
    if not Path(egp_path).exists():
        raise FileNotFoundError(f"not found {egp_path}")
    egp_path = Path(egp_path).resolve()
    return egp_path


def _open_enterprise_guide(eg_version: str) -> win32com.client.CDispatch:
    print(f"opening SAS Enterprise Guide {eg_version}")
    try:
        app = win32com.client.Dispatch(f"SASEGObjectModel.Application.{eg_version}")
    except pythoncom.com_error as e:
        if "Invalid class string" in str(e):
            raise SASEGVersionNotFound(
                f"The specified version {eg_version} cannot be found"
            )
        else:
            raise e
    click.secho("-> application instance created", fg="green")
    return app


def _activate_enterprise_guide_profile(
    profile_name: str, app: win32com.client.CDispatch
) -> None:
    print(f"activating profile:[{profile_name}]")
    try:
        app.SetActiveProfile(profile_name)
    except pythoncom.com_error as e:
        if "The given profile name does not exist" in str(e):
            raise SASEGProfileNotFound(
                f"The given profile name '{profile_name}' wasn't found in EG"
            )
        else:
            raise e
    click.secho(f"-> profile:[{profile_name}] activated", fg="green")


def _open_egp(egp_path: Path, app: win32com.client.CDispatch):
    print(f"opening {egp_path}")
    project_object = app.Open(str(egp_path), "")
    click.secho("-> egp file opened", fg="green")
    return project_object


def _run_egp(egp_path: str, project_object) -> None:
    print(f"running {egp_path}")
    project_object.Run()
    click.secho("-> run finished", fg="green")


def _check_log_for_errors(egp_path, log_dir):
    error_happend = False
    for log in log_dir.rglob("*.log"):
        with open(log, mode="r") as log_file:
            contents = log_file.read()
            error = re.search("^ERROR.*:", contents, re.MULTILINE)
            if error:
                click.secho(f"[{log.stem}] failed in {egp_path.name}", fg="red")
                start = error.start()
                last = contents.find("\n", start)
                print(contents[start:last])
                error_happend = True
                break
    return error_happend


def _save_and_close_egp(egp_path, overwrite, project_object):
    output = ""
    if overwrite:
        output = egp_path
    else:
        timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M")
        output = str(egp_path.parent / ("." + egp_path.stem + "_" + timestamp + ".egp"))

    project_object.SaveAs(output)
    click.secho(f"-> saved to {output}", fg="green")

    project_object.Close()
    return output


def _retrieve_logs(eg_version, verbose, output):
    print(f"getting logs from {output}")
    log_dir = Path(f"./{Path(output).stem}_CodeAndLogs")
    if log_dir.exists():
        shutil.rmtree(log_dir)
    res = subprocess.run(
        f'Cscript "{SCRIPTDIR_PATH}/ExtractCodeAndLog.vbs" "{output}" "{eg_version}"',
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=True,
    )
    if verbose:
        print(res.stdout)
    click.secho("-> log created", fg="green")
    return log_dir


def _finish_and_clean_up(egp_path, output, error_happend):
    if error_happend:
        os.rename(output, str(Path(output).parent / (Path(output).stem + ".ERROR.egp")))
        raise SASEGRuntimeError("Your EGP File had ERROR in the logs; please review")
    else:
        click.secho(f"successfully finished exectuing {egp_path.name}", fg="green")


def cli():
    """Cli interface for eg runner"""
    run_egp(sys.argv[1])


class SASEGRuntimeError(Exception):
    "Raised if the result logs of egp file include ERROR line"


class SASEGProfileNotFound(Exception):
    "Raised if the specified profile doesn't exist"


class SASEGVersionNotFound(Exception):
    "Raised if the specified version doesn't exist"


if __name__ == "__main__":
    # simple test
    # run_egp(SCRIPTDIR_PATH.parent / 'tests/test_fail.egp')
    run_egp(SCRIPTDIR_PATH.parent / "tests/test_success.egp")
