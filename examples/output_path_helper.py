from pathlib import Path


_EXAMPLES_DIR = Path(__file__).resolve().parent
_EXAMPLES_OUTPUT_DIR = _EXAMPLES_DIR / "outputfiles"


def examples_output_path(*parts: str) -> str:
    path = _EXAMPLES_OUTPUT_DIR.joinpath(*parts)
    path.parent.mkdir(parents=True, exist_ok=True)
    return str(path)


def ensure_examples_output_dir(*parts: str) -> str:
    path = _EXAMPLES_OUTPUT_DIR.joinpath(*parts)
    path.mkdir(parents=True, exist_ok=True)
    return str(path)
