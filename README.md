[![PyPI](https://img.shields.io/pypi/pyversions/saseg_runner.svg)](#)
[![PyPI](https://img.shields.io/pypi/status/saseg_runner.svg)](#)
[![PyPI](https://img.shields.io/pypi/v/saseg_runner)](https://pypi.org/project/pcm/)
[![PyPI](https://img.shields.io/pypi/l/saseg_runner.svg)](#)

# Overview
This repository gets you run egp file from python or command line.

# Install
```bash
pip install saseg_runner
```

# Usage

## as python library
```python
from saseg_runner import run_egp, SASEGRuntimeError

run_egp(egp_path='sample.egp', egp_version='7.1', profile_name='Your Profile')
# if some tasks in the egp file fails, this will raise SASEGRuntimeError.

"""
Other Parameters
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
```

## as command line
```bash
run_egp <your egp file path>
```


