import json
import subprocess
import sys

NOTEBOOK_PATH = "Asyncio test.ipynb"
PATCHED_PATH = "Asyncio test_patched.ipynb"

# Load the notebook
with open(NOTEBOOK_PATH, "r", encoding="utf-8") as f:
    nb = json.load(f)

# The import cell to prepend
import_cell = {
    "cell_type": "code",
    "execution_count": None,
    "metadata": {},
    "outputs": [],
    "source": [
        "import asyncio\n",
        "import aiohttp\n",
        "from comap import api_async\n",
        "from dotenv import dotenv_values"
    ]
}

# Session + auth cell to insert after the secrets cell
session_cell = {
    "cell_type": "code",
    "execution_count": None,
    "metadata": {},
    "outputs": [],
    "source": [
        "session = aiohttp.ClientSession()\n",
        "identity = api_async.Identity(session, secrets['COMAP_KEY'])\n",
        "token = await identity.authenticate(client_id, secret)\n",
        "print('Token obtained:', token is not None)"
    ]
}

# Insert import_cell at position 0 (before Markdown ## Initialization)
# and session_cell after the secrets cell (position 2)
nb["cells"].insert(0, import_cell)
# After insert, secrets cell is at index 2; insert session after it
nb["cells"].insert(3, session_cell)

# Save patched notebook
with open(PATCHED_PATH, "w", encoding="utf-8") as f:
    json.dump(nb, f, indent=1)

print(f"Patched notebook saved to: {PATCHED_PATH}")

# Run with nbconvert
result = subprocess.run(
    [
        sys.executable, "-m", "jupyter", "nbconvert",
        "--to", "notebook",
        "--execute",
        "--inplace",
        "--ExecutePreprocessor.timeout=60",
        PATCHED_PATH
    ],
    capture_output=True,
    text=True
)
print("STDOUT:", result.stdout)
print("STDERR:", result.stderr)
print("Exit code:", result.returncode)
