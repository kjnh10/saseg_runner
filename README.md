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
## as command line
```bash
run_egp <your egp file path>

# if some task in the project fails, 
```

## as python library
```python
from saseg_runner import run_egp, SASEGRuntimeError

run_egp(egp_path='sample.egp', egp_version='7.1', profile_name='Your Profile')
# if some tasks in the egp file fails, this will raise SASEGRuntimeError.
```
