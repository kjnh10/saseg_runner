from saseg_runner import run_egp


def test_fail():
    run_egp(SCRIPTDIR_PATH / 'test_fail.egp')

def test_success():
    run_egp(SCRIPTDIR_PATH / 'test_no_error.egp')
