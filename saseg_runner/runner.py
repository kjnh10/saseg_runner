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
import time
import datetime

SCRIPTDIR_PATH = Path(os.path.dirname(__file__)).resolve()


def run_egp(
        egp_path: Union[str, Path],
        profile_name: str = 'SAS Asia',
        eg_version: str = '7.1',
        overwrite: bool = False,
        remove_log: bool = True,
        verbose: bool = False,
        ) -> None:
    """
    execute egp_path

    Parameters
    ------------
    egp_path : Union[str, Path]
        SAS Enterprise Guide file path.
    profile_name : str
        profile name to use
    overwrite: bool
        controls whether to save the egp file after exection. if False, timestamp is added to filename. The default is False.
    remove_log: bool
        wether remove log files or not. the default is True.
    verbose: bool
        default is False
    """
    start_time = time.time()

    if not Path(egp_path).exists():
        raise Exception(f'not found {egp_path}')
    egp_path = Path(os.path.abspath(egp_path))

    print(f'opening SAS Enterprise Guide {eg_version}')
    app = win32com.client.Dispatch(f'SASEGObjectModel.Application.{eg_version}')
    click.secho('-> application instance created', fg='green')

    print(f'activating profile:[{profile_name}]')
    app.SetActiveProfile(profile_name)
    click.secho(f'-> profile:[{profile_name}] activated', fg='green')

    print(f'opening {egp_path}')
    prjObject = app.Open(str(egp_path), "")
    click.secho('-> egp file opened', fg='green')

    print(f'running {egp_path}')
    prjObject.Run()
    click.secho('-> run finished', fg='green')

    output = ""
    if overwrite:
        output = egp_path
    else:
        timestamp = datetime.datetime.now().strftime('%Y%m%d-%H%M')
        output = str(egp_path.parent / ('.' + egp_path.stem + '_' + timestamp + '.egp'))

    prjObject.SaveAs(output)
    click.secho(f'-> saved to {output}', fg='green')

    prjObject.Close()

    print(f'getting logs from {output}')
    log_dir = Path(f'./{Path(output).stem}_CodeAndLogs')
    if (log_dir.exists()):
        shutil.rmtree(log_dir)
    res = subprocess.run(
        f'Cscript "{SCRIPTDIR_PATH}/ExtractCodeAndLog.vbs" "{output}" "{eg_version}"',
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=True,
    )
    if verbose:
        print(res.stdout)
    click.secho('-> log created', fg='green')

    error_happend = False
    for log in log_dir.rglob('*.log'):
        with open(log, mode='r') as f:
            contents = f.read()
            error = re.search("^ERROR.*:", contents, re.MULTILINE)
            if (error):
                click.secho(f"[{log.stem}] failed in {egp_path.name}", fg="red")
                start = error.start()
                last = contents.find('\n', start)
                print(contents[start:last])
                error_happend = True
                break

    if remove_log:
        shutil.rmtree(log_dir)

    if error_happend:
        os.rename(output, str(Path(output).parent/(Path(output).stem+'.ERROR.egp')))
        raise SASEGRuntimeError
    else:
        click.secho(f'successfully finished exectuing {egp_path.name}', fg='green')

    elapsed_time = int(time.time() - start_time)
    print(f"elapsed_time:{elapsed_time}[sec]")


def cli():
    run_egp(sys.argv[1])


class SASEGRuntimeError(Exception):
    "error the result logs of egp file include ERROR line"


if __name__ == "__main__":
    # simple test
    # run_egp(SCRIPTDIR_PATH.parent / 'tests/test_fail.egp')
    run_egp(SCRIPTDIR_PATH.parent / 'tests/test_success.egp')
