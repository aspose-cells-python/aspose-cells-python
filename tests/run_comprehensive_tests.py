"""
Comprehensive Test Runner
Run all test categories including the new comprehensive tests for 100% coverage.
"""

import subprocess
import sys
from pathlib import Path


def run_comprehensive_tests():
    """Run the comprehensive test suite for 100% coverage."""
    
    project_root = Path(__file__).parent.parent
    tests_dir = Path(__file__).parent
    
    print("🚀 Running Comprehensive Excel Test Suite for 100% Coverage")
    print("=" * 60)
    
    # All test files in execution order
    test_files = [
        "test_units.py",                    # Core unit tests
        "test_excel_generation.py",        # Excel file generation
        "test_conversions.py",              # Basic format conversions
        "test_markitdown_plugin.py",       # MarkItDown plugin tests
        "test_advanced_converters.py",     # Advanced converter tests
        "test_io_modules.py",               # IO module tests
        "test_range_and_style.py",         # Range and style tests
        "test_utilities_and_edge_cases.py" # Utilities and edge cases
    ]
    
    total_passed = 0
    total_failed = 0
    
    for test_file in test_files:
        test_path = tests_dir / test_file
        if not test_path.exists():
            print(f"⚠️  Skipping {test_file} - file not found")
            continue
            
        print(f"\n📋 Running {test_file}...")
        print("-" * 40)
        
        result = subprocess.run([
            sys.executable, "-m", "pytest", 
            str(test_path),
            "-v",
            "--tb=short"
        ], cwd=project_root, capture_output=False)
        
        if result.returncode == 0:
            print(f"✅ {test_file} - PASSED")
            total_passed += 1
        else:
            print(f"❌ {test_file} - FAILED")
            total_failed += 1
    
    print("\n" + "=" * 60)
    print(f"📊 Comprehensive Test Summary:")
    print(f"   Passed: {total_passed}")
    print(f"   Failed: {total_failed}")
    print(f"   Total:  {total_passed + total_failed}")
    
    if total_failed == 0:
        print("🎉 All comprehensive test categories passed!")
        return 0
    else:
        print("💥 Some test categories failed!")
        return 1


def run_with_coverage():
    """Run comprehensive tests with coverage reporting."""
    
    project_root = Path(__file__).parent.parent
    
    print("📊 Running Comprehensive Tests with Coverage...")
    
    # Run all test files
    test_files = [
        "tests/test_units.py",
        "tests/test_excel_generation.py", 
        "tests/test_conversions.py",
        "tests/test_markitdown_plugin.py",
        "tests/test_advanced_converters.py",
        "tests/test_io_modules.py",
        "tests/test_range_and_style.py",
        "tests/test_utilities_and_edge_cases.py"
    ]
    
    result = subprocess.run([
        sys.executable, "-m", "pytest"
    ] + test_files + [
        "-v",
        "--cov=aspose",
        "--cov-report=html",
        "--cov-report=term-missing",
        "--cov-report=xml",
        "--cov-fail-under=95",
        "--tb=short"
    ], cwd=project_root)
    
    return result.returncode


def run_quick_check():
    """Run a quick check of all test files."""
    
    project_root = Path(__file__).parent.parent
    
    print("⚡ Running Quick Test Check...")
    
    result = subprocess.run([
        sys.executable, "-m", "pytest", 
        "tests/",
        "-q",
        "--tb=no"
    ], cwd=project_root)
    
    if result.returncode == 0:
        print("✅ Quick check passed!")
    else:
        print("❌ Quick check failed!")
    
    return result.returncode


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Run comprehensive test suite")
    parser.add_argument("--coverage", action="store_true", 
                       help="Run with coverage reporting")
    parser.add_argument("--quick", action="store_true",
                       help="Run quick check only")
    
    args = parser.parse_args()
    
    if args.quick:
        exit_code = run_quick_check()
    elif args.coverage:
        exit_code = run_with_coverage()
    else:
        exit_code = run_comprehensive_tests()
    
    sys.exit(exit_code)
