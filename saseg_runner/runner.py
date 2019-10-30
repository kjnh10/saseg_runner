# %%
import win32com.client
import subprocess
import re
import os
from pathlib import Path
import shutil
from typing import Union
import sys
import click

SCRIPTDIR_PATH = Path(os.path.dirname(__file__)).resolve()


def run_egp(
        egp_path: Union[str, Path],
        eg_version: str = '7.1',
        profile_name: str = 'SAS Asia',
        remove_log: bool = True,
        ) -> None:
    """
    execute egp_path
    return True if execution log has no error log.
    """
    if not Path(egp_path).exists():
        raise Exception(f'not found {egp_path}')
    egp_path = Path(egp_path).resolve()

    app = win32com.client.Dispatch(f'SASEGObjectModel.Application.{eg_version}')
    app.SetActiveProfile(profile_name)
    prjObject = app.Open(str(egp_path), "")
    prjObject.Run()
    click.secho('run finished', fg='green')
    prjObject.Save()
    click.secho('saved', fg='green')
    prjObject.Close()

    log_dir = Path(f'./{egp_path.stem}_CodeAndLogs')
    if (log_dir.exists()):
        shutil.rmtree(log_dir)

    subprocess.run(f'Cscript {SCRIPTDIR_PATH}/ExtractCodeAndLog.vbs {egp_path} {eg_version}')
    print('log created')

    error_happend = False
    for log in log_dir.rglob('*.log'):
        with open(log, mode='r') as f:
            contents = f.read()
            error = re.search("^ERROR:", contents, re.MULTILINE)
            if (error):
                click.secho(f"[{log.stem}] failed in {egp_path.name}", fg="red")
                start = error.start()
                last = contents.find('\n', start)
                print(contents[start:last])
                error_happend = True
                break
    if error_happend:
        if remove_log:
            shutil.rmtree(log_dir)
        raise SASEGRuntimeERROR
    else:
        if remove_log:
            shutil.rmtree(log_dir)
        click.secho(f'successfully finished exectuing {egp_path.name}', fg='green')


def cli():
    run_egp(sys.argv[1])


class SASEGRuntimeERROR(Exception):
    "error the result logs of egp file include ERROR line"


if __name__ == "__main__":
    # simple test
    run_egp(SCRIPTDIR_PATH / 'test_fail.egp')
    # run_egp(SCRIPTDIR_PATH / 'test_no_error.egp')
