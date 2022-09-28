r"""
Source environment runner for direct no-install debugging & testing.
Includes setup tool for mini-venv '.senv/'

Usage:

  python senv.py [OPTIONS] [PY_SCRIPT] [SCRIPTARGSOPTS]

  PY_SCRIPT: default win32/test/testall.py
  OPTIONS:   see below

Examples - direct usage (does post-mortem debugging by default):

  python senv.py pywin32_testall.py -v -skip-adodbapi
  python -u senv.py -c com/win32com/test/testall.py -v 1
  python senv.py -D SOME_SCRIPT_OR_TEST.py    # debug in pywin.debugger
  python[w] senv.py -W -new                   # starts Pythonwin

Examples - usage via the virtual mini-venv '.senv/':

  # Finish setup of the venv  (Once for every python version to use)
  python senv.py -X    # creates the venv python[NN].exe links for that python

  # Activate & use
  .senv\Scripts\activate
  python win32/test/testall.py -v
  python                             # interactive
  python37                           # use particular python version
  python Pythonwin/start_pythonwin.pyw -new
  ...
  deactivate

  # Use virtual env once without further activation
  .senv\Scripts\python37 pywin32_testall.py -v -skip-adodbapi

  # Note: In virtual .senv subprocess.Popen is not (and doesn't need to be)
  # hooked unlike w direct senv.py usage. Complex tests with installations /
  # registrations relying on sys.executable / sys.exec_prefix (Services, COM
  # servers, ..)  may use the source environment consistently only via the
  # virtual .senv method.

Runs a python (test) script inplace using pywin32 "virtually" directly from
this source tree and (optional) build_ext products instead of requiring an
installation. Enables a quicker, more direct (iterative) debugging & testing
experience. Works by fixing up sys.path, sys.modules, environ['PATH'],
subprocess.Popen, some package __path__'s, etc. Options for trace & post-mortem
debugging. Does post-mortem debugging in Pythonwin by default.
"""

OPTIONS = r"""
OPTIONS:
        elif a == "-c":
            chdir_to_scriptdir = 1  # chdir to script dir
        elif a == "-d":
            do_trace = 1  # debug it in pywin.debugger
        elif a == "-D":
            do_trace = 2  # debug it in pdb
        elif a == "-i":
            script = __import__("code").__file__.replace(".pyc", ".py")  # interactive
            break
        elif a == "-m":
            script = __import__("runpy").__file__.replace(".pyc", ".py")  # run module
            break
        elif a == "-P":
            debug_pm = 0  # no PM debugging (default pywin.debugger PM)
        elif a == "-p":
            debug_pm = 2  # pdb PM debugging
        elif a == "-V":
            verbose = (verbose - 1) % 3  # toggle verbosity
        elif a == "-W":
            script = join(__filedir__, "Pythonwin/start_pythonwin.pyw")
            break
        elif a in ("-X", "--setup-env"):
            sys.exit(setup_venv())  # ensure python symlinks in '.senv/'"""

import sys, os, re, types
import sysconfig
import subprocess

join = os.path.join

__fileabs__ = __file__ = os.path.abspath(__file__)
__filedir__ = os.path.dirname(__fileabs__)


# Called by .senv/sitecustomize.py .
def use_source_env(patch_popen=False):
    run_in_srctree(None, patch_popen=patch_popen)
    if verbose:
        print("-- Using source env with sys.path=%r" % (sys.path,))


def run_in_srctree(script="win32/test/testall.py", patch_popen=True):
    if verbose > 1:
        print("-- sys.modules before distutils import: %r" % modfiles())
    if hasattr(sys.stdout, "errors") and sys.stdout.errors == "surrogateescape":
        sys.stdout.reconfigure(errors="backslashreplace")

    sp_old = sys.path[:]
    envpath_old = os.environ["PATH"]
    sa_old = sys.argv[:]
    main_old = sys.modules["__main__"]
    cwd = os.getcwd()
    _unroll = True
    try:
        if verbose > 1:
            print("-- sys.modules before inplace patch: %r" % modfiles())

        # Drop conflicting environment
        ps32 = sys.modules.pop("pywin32_system32", None)
        if ps32:
            if verbose > 1:
                print("-- drop pywin32_system32: %r" % ps32.__path__)
            ps32.__path__[:] = []
        pb = sys.modules.pop("pywin32_bootstrap", None)
        if pb and hasattr(os, "add_dll_directory"):
            if hasattr(pb, "h_add_dll"):
                # close old search path for pywintypesNN.dll
                pb.h_add_dll.close()
            elif verbose:
                print("-- CANNOT CLOSE %s.h_add_dll!  Check pywintypes path!" % pb)
        for p in os.environ["PATH"].split(os.pathsep):
            if p.endswith("\\pywin32_system32"):
                os.environ["PATH"] = os.environ["PATH"].replace(p + os.pathsep, "")
        wp_x_removes = []
        for p in sys.path[:]:
            for x in (
                "\\win32",
                "\\win32\\lib",
                "\\Pythonwin",
            ):
                if p.endswith(x):
                    sys.path.remove(p)
                    if not p.endswith("\\win32\\lib"):
                        # path will potentially be re-added in mixed env
                        wp_x_removes.append(p)

        # Add new environment: Source folders for *.py, build_platlib folders for extensions
        for p in [
            "Pythonwin",
            "com",
            "win32/Lib",
        ]:
            sys.path.insert(0, join(__filedir__, p))

        bplatlib_abs = get_build_platlib()
        bplatlib_exists = []
        for p in [
            "Pythonwin",
            "win32",
            "",  # -> import pywin32_system32
        ]:
            p = join(bplatlib_abs, p)
            sys.path.insert(0, p)
            if os.path.isdir(p):
                bplatlib_exists.extend(os.listdir(p))
        if len(bplatlib_exists) < 20:
            if verbose:
                print(
                    "-- Only %d build_ext products found. Retaining for mixed import in sys.path: %s"
                    % (len(bplatlib_exists), wp_x_removes),
                )
            sys.path.extend(wp_x_removes)
        else:
            wp_x_removes = None

        # Prepend ./~SENVPATH1 to sys.path to enable custom import overrides
        sys.path.insert(0, join(__filedir__, "~SENVPATH1"))

        # import pywin32_bootstrap again for pywintypesNN.dll / pythoncomNN.dll from build_ext (potentially)
        import pywin32_system32  # noqa "unused"
        import pywin32_bootstrap  # noqa "unused"
        import pywintypes
        import win32com

        # Link win32comext pyd's from build tree
        def win32com__PackageSupportBuildPath__(__path__):
            m = re.search(r"com[/\\]win32comext[/\\](\w+)$", __path__[0])
            if m:
                __path__.append(join(bplatlib_abs, "win32comext", m.group(1)))
                if wp_x_removes:
                    import site

                    # mixed env
                    __path__.append(
                        join(site.getsitepackages()[-1], "win32comext", m.group(1))
                    )

        win32com.__PackageSupportBuildPath__ = win32com__PackageSupportBuildPath__
        if verbose:
            print("-- pywintypes found at %r" % pywintypes.__file__)

        if patch_popen:
            patch_Popen()

        if not script:
            # Ready for use_source_env()
            _unroll = False
            if verbose:
                print()
            return

        # Script & args
        script_abs = os.path.abspath(script)
        scriptdir, script_bn = os.path.split(script_abs)
        sys.path.insert(1, scriptdir)  # insert behind SENVPATH1
        sys.argv[:] = [script_abs] + sys.argv[1:]

        # A clean separate __main__ module
        main_exec = types.ModuleType("__main__")
        main_exec.__file__ = script_abs
        sys.modules["__main__"] = main_exec

        if verbose:
            if verbose < 2:
                print("-- run %r in source env %r" % (script, __filedir__))
            else:
                print("-- running %r with sys.path=%r" % (script, sys.path))
                print("--     and PATH=%r" % os.environ["PATH"])
        scr = open(script_abs, "rb").read()
        co = compile(scr, script_abs, "exec", dont_inherit=True)
        if chdir_to_scriptdir and scriptdir != cwd:
            if verbose:
                print("-- chdir to %r" % (scriptdir,))
            os.chdir(scriptdir)
        if verbose:
            print()
            sys.stdout.flush()
        if do_trace == 1:
            import pywin.debugger

            pywin.debugger.set_trace()
        elif do_trace == 2:
            import pdb

            pdb.set_trace()

        exec(co, main_exec.__dict__)  # enter into script. pwdb: press F11

    except Exception:
        if not debug_pm:
            _unroll = False
            raise
        import traceback

        traceback.print_exc()
        sys.stdout.flush()
        sys.stderr.flush()
        if debug_pm == 1:
            try:
                import pywin.debugger

                pywin.debugger.post_mortem()
                return
            except Exception:
                pass  # fall back to pdb
        import pdb

        pdb.post_mortem()
    finally:
        if _unroll:
            # Unroll patches as far as possible
            if os.getcwd() != cwd:
                os.chdir(cwd)
            sys.modules["__main__"] = main_old
            sys.argv[:] = sa_old
            unpatch_Popen()
            if 0:  # We don't really get rid of imported DLL's/pyd's ...
                pbn = sys.modules.pop("pywin32_bootstrap", None)
                if pbn:
                    # close the new search path for pywintypesNN.dll though it might not help w/o unloading all the bound dlls ...
                    pb.dll_add_handle.close()
                if pb:
                    sys.modules["pywin32_bootstrap"] = pb
            os.environ["PATH"] = envpath_old
            sys.path[:] = sp_old
            if verbose:
                print("-- end of source env run")


def get_build_platlib():
    # Platlib build dir new format, e.g. "build\\lib.win-amd64-cpython-37"
    bplatlib = "build\\lib.%s-%s" % (
        sysconfig.get_platform(),
        hasattr(sys, "implementation")
        and sys.implementation.cache_tag
        or "cpython-%d%d" % sys.version_info[:2],
    )
    bplatlib_abs = bplatlib_new = join(__filedir__, bplatlib)
    if os.path.isdir(bplatlib_abs):
        return bplatlib_abs

    # Check old format e.g. "build\\lib.win-amd64-3.10"
    bplatlib = "build\\lib.%s-%s" % (
        sysconfig.get_platform(),
        sysconfig.get_python_version(),
    )
    bplatlib_abs = join(__filedir__, bplatlib)
    if os.path.isdir(bplatlib_abs):
        return bplatlib_abs

    if verbose:
        print("-- platlib build dir doesn't exist yet!: %s" % bplatlib_new)
    return bplatlib_new

    raise NotImplementedError
    # Ask setuptools/distutils - A massive slow import unfortunately

    # XXX setuptools/pkg_resources (v65.1.0/21.3) tries to import win32com.shell
    # first !? (at pkg_resources._vendor.appdirs:560). Possibly replace that import
    # line with "raise ImportError" to let use _get_win_folder_with_ctypes() which
    # should be default propbably.
    # It is unclear if `import win32com.shell` failes "well" without messing our
    # pristine import / __path__ state with conflicting stuff at this stage before
    # we actually add the new platlib paths for win32* as primary paths.

    import warnings
    import setuptools
    import distutils.command.build

    build = distutils.command.build.build(setuptools.Distribution())
    build.ensure_finalized()
    bplatlib = build.build_platlib  # 'build\\lib.win-amd64-cpython-37'
    bplatlib_abs = join(__filedir__, bplatlib)
    if bplatlib_abs != bplatlib_new:
        warnings.warn(
            "-- UNEXPECTED platlib build dir! CONSIDER upgrading setuptools or adjusting this code"
        )
    if verbose and not os.path.isdir(bplatlib_abs):
        print("-- doesn't exist!: %s" % bplatlib_abs)

    return bplatlib_abs


def patch_Popen():
    if subprocess.Popen.__module__ == "subprocess":
        # patch Popen (used by pywin32_testall.py and win32/test/testall.py for win32/Demos/*) to use this
        # same source env runner instead of bare sys.executable
        global _Popen_orig
        _Popen_orig = subprocess.Popen

        class PopenPatch(_Popen_orig):
            def __init__(self, args, *rest, **kw):
                if isinstance(args, (tuple, list)):
                    prog = args[0]
                    if prog == sys.executable:
                        import itertools

                        popts = list(
                            itertools.takewhile(
                                lambda opt: opt.startswith("-"), args[1:]
                            )
                        )
                        args = (
                            [sys.executable]
                            + popts
                            + [__fileabs__, "-P"]
                            + list(args[1 + len(popts) :])
                        )
                        if verbose:
                            print("-- PopenPatch: %r %r %r" % (args, rest, kw))
                elif sys.executable in args:
                    raise NotImplementedError("TODO code for string cmd arg")
                _Popen_orig.__init__(self, args, *rest, **kw)

        subprocess.Popen = PopenPatch


def unpatch_Popen():
    if subprocess.Popen.__module__ != "subprocess":
        subprocess.Popen = _Popen_orig


def modfiles():
    return [getattr(mod, "__file__", name) for name, mod in sys.modules.items()]


# Ensures the mini-venv setup up for the running python version by ensuring the
# up-to-date symlinks to that python executables.
def setup_venv(env_dir=".senv", sim=False):
    import venv

    v = venv.EnvBuilder()
    v.symlinks = True
    c = types.SimpleNamespace()
    # src
    c.executable = sys._base_executable
    if "python.exe" not in c.executable.lower():
        c.executable = join(sys.exec_prefix, "python.exe")
    c.python_dir, c.exename = os.path.split(c.executable)
    # dst
    c.env_dir = env_dir
    c.bin_path = join(__filedir__, c.env_dir, "Scripts")

    # Link like v.setup_python(c)
    exes = [
        (f, f)
        for f in os.listdir(c.python_dir)
        if f.endswith(".exe") and "python" in f or f.endswith(".dll")
    ] + [
        ("python.exe", "python%s.exe" % sys.winver.replace(".", "")),
        ("pythonw.exe", "pythonw%s.exe" % sys.winver.replace(".", "")),
    ]
    for exe, exe_t in exes:
        src = os.path.join(c.python_dir, exe)
        if os.path.lexists(src):
            dst = os.path.join(c.bin_path, exe_t)
            if os.path.exists(dst):
                if os.path.islink(dst) and os.readlink(dst) == src:
                    print("-- symlink was ok : %s -> %s" % (dst, src))
                    continue
                print("-- remove: %s" % (dst,))
                if not sim:
                    os.remove(dst)
            print("-- make symlink: %s -> %s" % (dst, src))
            if not sim:
                try:
                    os.symlink(src, dst)
                except OSError as ev:
                    if "privilege" not in str(ev):
                        raise
                    print("--   fall back to hardlink/copy")
                    try:
                        os.link(src, dst)
                    except OSError:
                        import shutil

                        shutil.copy2(src, dst)


def usage(opt):
    if opt == "-h":
        print(
            r"""\
Usage:

  python senv.py [OPTIONS] [PY_SCRIPT] [SCRIPTARGSOPTS]

  PY_SCRIPT: default win32/test/testall.py
"""
            + OPTIONS
        )
    else:
        print(__doc__ + OPTIONS)


verbose = not (not sys.stdout or not sys.stdout.encoding)  # 0 when pythonw.exe
do_trace = 0  # 1 = pywin.debugger; 2 = pdb
debug_pm = 1  # 1 = pywin.debugger; 2 = pdb; 0 = None/raise
chdir_to_scriptdir = 0

if __name__ == "__main__":
    i = 1
    script = "win32/test/testall.py"
    while sys.argv[i:]:
        a = sys.argv[i]
        i += 1
        if a in ("-h", "--help"):
            sys.exit(usage(a))
        elif a == "-c":
            chdir_to_scriptdir = 1  # chdir to script dir
        elif a == "-d":
            do_trace = 1  # debug it in pywin.debugger
        elif a == "-D":
            do_trace = 2  # debug it in pdb
        elif a == "-i":
            script = __import__("code").__file__.replace(".pyc", ".py")  # interactive
            break
        elif a == "-m":
            script = __import__("runpy").__file__.replace(".pyc", ".py")  # run module
            break
        elif a == "-P":
            debug_pm = 0  # no PM debugging (default pywin.debugger PM)
        elif a == "-p":
            debug_pm = 2  # pdb PM debugging
        elif a == "-V":
            verbose = (verbose - 1) % 3  # toggle verbosity
        elif a == "-W":
            script = join(__filedir__, "Pythonwin/start_pythonwin.pyw")
            break
        elif a in ("-X", "--setup-env"):
            sys.exit(setup_venv())  # ensure python symlinks in '.senv/'
        elif a == ("--"):
            break  # end option parsing, forward remainder to script
        elif a.startswith("-"):
            raise RuntimeError("unkown option %r", a)
        else:
            # non-default script, forward remainder args
            script = a
            break

    sa_old = sys.argv
    try:
        sys.argv[1:] = sys.argv[i:]
        run_in_srctree(script)
    finally:
        sys.argv[:] = sa_old
