# %%
import win32com.client
import win32gui
import subprocess
import re
from pathlib import Path
import shutil
from typing import Union
import click

def run_egp(egp_path: Union[str, Path], eg_version: str='7.1', profile_name='SAS Asia') -> None:
    """
    execute egp_path
    return True if execution log has no error log.
    """
    app = win32com.client.Dispatch(f'SASEGObjectModel.Application.{eg_version}')
    app.SetActiveProfile(profile_name)
    prjObject = app.Open(str(egp_path), "")
    prjObject.Run()
    click.secho('run finished', fg='green')
    prjObject.Save()
    click.secho('saved', fg='green')
    # prjObject.Close()

    log_dir = Path(f'./{egp_path.stem}_CodeAndLogs')
    if (log_dir.exists()):
        shutil.rmtree(log_dir)

    subprocess.run(f'Cscript ExtractCodeAndLog.vbs {egp_path.name}')
    print('log created')

    for log in log_dir.rglob('*.log'):
        with open(log, mode='r') as f:
            contents = f.read()
            error = re.search("^ERROR:", contents, re.MULTILINE)
            if (error):
                click.secho(f"[{log.stem}] failed in {egp_path.name}", fg="red")
                start = error.start()
                last = contents.find('\n', start)
                print(contents[start:last])
                raise SASEGRuntimeERROR

    click.secho(f'successfully finished exectuing {egp_path.name}', fg='green')


class SASEGRuntimeERROR(Exception):
    "error the result logs of egp file include ERROR line"


if __name__ == "__main__":
    # run_egp(Path('./test_fail.egp'))
    run_egp(Path('./test_no_error.egp'))
