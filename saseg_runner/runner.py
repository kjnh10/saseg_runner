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
import win32com.client

SCRIPTDIR_PATH = Path(__file__).parent.resolve()
DEFAULT_PROFILE_NAME = "SAS Asia"
DEFAULT_EG_VERSION = "7.1"


def run_egp(
    egp_path: Union[str, Path],
    profile_name: str = DEFAULT_PROFILE_NAME,
    eg_version: str = DEFAULT_EG_VERSION,
    overwrite: bool = False,
    remove_log: bool = True,
    verbose: bool = False,
) -> None:
    """[summary]

    Args:
        egp_path (Union[str, Path]): SAS Enterprise Guide file path.
        profile_name (str, optional): profile name to use. Defaults to 'SAS Asia'.
        eg_version (str, optional): Which version of EG to use. Defaults to '7.1'.
        overwrite (bool, optional): controls whether to save the egp file after exection. if False, timestamp is added to filename. Defaults to False.
        remove_log (bool, optional): Whether to remove log files or not. Defaults to True.
        verbose (bool, optional): [description]. Defaults to False.
    """
    start_time = time.time()

    egp_path = _check_egp_file_existence(egp_path)
    app = _open_enterprise_guide(eg_version)
    _activate_enterprise_guide_profile(profile_name, app)
    prjObject = _open_egp(egp_path, app)
    _run_egp(egp_path, prjObject)
    output = _save_and_close_egp(egp_path, overwrite, prjObject)
    log_dir = _retrieve_logs(eg_version, verbose, output)
    error_happend = _check_log_for_errors(egp_path, log_dir)
    if remove_log:
        shutil.rmtree(log_dir)
    _finish_and_clean_up(egp_path, output, error_happend)
    elapsed_time = int(time.time() - start_time)
    print(f"elapsed_time:{elapsed_time}[sec]")


def _check_egp_file_existence(egp_path: str) -> Path:
    if not Path(egp_path).exists():
        raise Exception(f"not found {egp_path}")
    egp_path = Path(egp_path).resolve()
    return egp_path


def _open_enterprise_guide(eg_version: str) -> win32com.client.CDispatch:
    print(f"opening SAS Enterprise Guide {eg_version}")
    app = win32com.client.Dispatch(f"SASEGObjectModel.Application.{eg_version}")
    click.secho("-> application instance created", fg="green")
    return app


def _activate_enterprise_guide_profile(
    profile_name: str, app: win32com.client.CDispatch
) -> None:
    print(f"activating profile:[{profile_name}]")
    app.SetActiveProfile(profile_name)
    click.secho(f"-> profile:[{profile_name}] activated", fg="green")


def _open_egp(egp_path: Path, app: win32com.client.CDispatch):
    print(f"opening {egp_path}")
    prjObject = app.Open(str(egp_path), "")
    click.secho("-> egp file opened", fg="green")
    return prjObject


def _run_egp(egp_path: str, prjObject) -> None:
    print(f"running {egp_path}")
    prjObject.Run()
    click.secho("-> run finished", fg="green")


def _check_log_for_errors(egp_path, log_dir):
    error_happend = False
    for log in log_dir.rglob("*.log"):
        with open(log, mode="r") as f:
            contents = f.read()
            error = re.search("^ERROR.*:", contents, re.MULTILINE)
            if error:
                click.secho(f"[{log.stem}] failed in {egp_path.name}", fg="red")
                start = error.start()
                last = contents.find("\n", start)
                print(contents[start:last])
                error_happend = True
                break
    return error_happend


def _save_and_close_egp(egp_path, overwrite, prjObject):
    output = ""
    if overwrite:
        output = egp_path
    else:
        timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M")
        output = str(egp_path.parent / ("." + egp_path.stem + "_" + timestamp + ".egp"))

    prjObject.SaveAs(output)
    click.secho(f"-> saved to {output}", fg="green")

    prjObject.Close()
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
    run_egp(sys.argv[1])


class SASEGRuntimeError(Exception):
    "Raised if the result logs of egp file include ERROR line"


if __name__ == "__main__":
    # simple test
    # run_egp(SCRIPTDIR_PATH.parent / 'tests/test_fail.egp')
    run_egp(SCRIPTDIR_PATH.parent / "tests/test_success.egp")
