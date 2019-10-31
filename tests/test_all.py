import subprocess
import os
import pytest
from pathlib import Path
from saseg_runner import run_egp
from saseg_runner.runner import SASEGRuntimeError
SCRIPTDIR = Path(os.path.dirname(__file__)).resolve()

fail_egp = SCRIPTDIR / 'test_fail.egp'
success_egp = SCRIPTDIR / 'test_success.egp'


def test_success():
    run_egp(str(success_egp))


def test_success_with_Path_object():
    run_egp(success_egp)


def test_fail():
    with pytest.raises(SASEGRuntimeError):
        run_egp(str(fail_egp))


def test_cli_success():
    res = subprocess.run(f'run_egp {str(success_egp)}')
    assert(res.returncode == 0)
