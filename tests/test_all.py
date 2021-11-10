import subprocess
from pathlib import Path

import pytest
import saseg_runner as sas


SCRIPTDIR = Path(__file__).parent.resolve()
PROFILE_NAME = "SAS Asia"

fail_egp = SCRIPTDIR / "test_fail.egp"
success_egp = SCRIPTDIR / "test_success.egp"
japanese_char_in_path_egp = SCRIPTDIR / "日本語フォルダ名/test_success.egp"
space_in_path_egp = SCRIPTDIR / "space exist/test_success.egp"


def test_success():
    sas.run_egp(str(success_egp), PROFILE_NAME)


def test_success_with_Path_object():
    sas.run_egp(success_egp, PROFILE_NAME)


def test_japanese_char_in_path():
    sas.run_egp(str(japanese_char_in_path_egp), PROFILE_NAME)


def test_space_in_path():
    sas.run_egp(str(space_in_path_egp), PROFILE_NAME)


def test_fail():
    with pytest.raises(sas.SASEGRuntimeError):
        sas.run_egp(str(fail_egp), PROFILE_NAME)


def test_fail_error_format2():  # there exists an error format like "ERROR 22-232: ......."
    with pytest.raises(sas.SASEGRuntimeError):
        sas.run_egp(str(SCRIPTDIR / "test_error_format2.egp"), PROFILE_NAME)


def test_cli_success():
    res = subprocess.run(f"run_egp {str(success_egp)}", PROFILE_NAME)
    assert res.returncode == 0


def test_eg_runner():
    eg = sas.EGRunner(PROFILE_NAME)
    eg.run_egp(success_egp)
    eg.run_egp(success_egp)