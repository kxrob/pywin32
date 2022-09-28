try:
    import sys, os

    srcdir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if srcdir not in sys.path:
        sys.path.insert(0, srcdir)

    import senv

    ##senv.verbose = 0
    senv.use_source_env()

except Exception as ev:
    raise BaseException(*ev.args)  # break vs. warn in python startup
