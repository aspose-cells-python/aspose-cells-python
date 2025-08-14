"""
Simple Test Runner - Run simplified tests quickly
"""

import subprocess
import sys
from pathlib import Path


def run_tests():
    """Run simplified tests."""
    
    project_root = Path(__file__).parent.parent
    
    print("🚀 Running Simplified Test Suite")
    print("=" * 40)
    
    # Run tests with basic coverage
    cmd = [
        sys.executable, "-m", "pytest", 
        "tests/",
        "-v",
        "--tb=short",
        "-x"  # Stop on first failure
    ]
    
    result = subprocess.run(cmd, cwd=project_root)
    
    if result.returncode == 0:
        print("\n✅ All simplified tests passed!")
        
        # Run coverage check
        print("\n📊 Checking coverage...")
        coverage_cmd = [
            sys.executable, "-m", "pytest",
            "tests/",
            "--cov=aspose",
            "--cov-report=term-missing",
            "--cov-report=html",
            "-q"
        ]
        
        coverage_result = subprocess.run(coverage_cmd, cwd=project_root)
        return coverage_result.returncode
    else:
        print("\n❌ Some tests failed!")
        return result.returncode


if __name__ == "__main__":
    exit_code = run_tests()
    sys.exit(exit_code)
