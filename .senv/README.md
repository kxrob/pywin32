# .senv/ - Venv for direct no-install debugging & testing

This is a source environment runner
mini-[`venv`](https://docs.python.org/library/venv.html) which makes Python
(tests, scripts, interactive, Pythonwin, debugger ...) use `pywin32` directly
inplace from this source tree and (optional) `build_ext` products, instead of
requiring an installation.

Note: The source env runner can also be used ad-hoc without this venv via
([../senv.py](../senv.py)).

## Finish setup of '.senv/'
(Once for any python version to use)

    python senv.py -X    # creates the venv python[NN].exe links for that python

## Activate & use
    .senv\Scripts\activate
    python win32/test/testall.py -v    # uses last python version set up
    python                             # interactive
    python37                           # particular python version
    python Pythonwin/start_pythonwin.pyw -new
    ...
    deactivate

## Use virtual env once without further activation
    .senv\Scripts\python37 pywin32_testall.py -v -skip-adodbapi

## Notes

When there are only few (or no) `build_ext` products (C-extensions in the
`build/` folder), the environment is overlayed to the potentially installed
pywin32. This allows to quickly test small changes to the C++ code (e.g. after
`python setup.py build_sel <MODULENAME>...`) or Python only changes. The
installed pywin32 then should be a more recent build of course.

## See also

Example for using the source env runner directly without the venv:

    python senv.py -D SOME_SCRIPT_OR_TEST.py    # debug in pywin.debugger

For further documentation see [../senv.py](../senv.py).
