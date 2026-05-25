import os, shutil

PROJECT_TMP = os.path.join(os.path.dirname(__file__), "..", ".test_tmp")


def pytest_sessionfinish(session):
    """Clean up .test_tmp after all tests finish (visible proof nothing leaks)."""
    if os.path.isdir(PROJECT_TMP):
        try:
            shutil.rmtree(PROJECT_TMP)
        except OSError:
            pass
