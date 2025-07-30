from cx_Freeze import setup, Executable
import sys
import os

# Add the parent directory to the Python path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from version_manager.version_manager import get_current_version

sys.setrecursionlimit(1500)

setup(
    name="UltimateTransactionCompiler",
    version=get_current_version(),
    description="Ultimate version of Transaction Compiler",
    executables=[Executable("main.py")]
)

