"""Run the node harness that exercises the wizard's mandatory-field mapping."""
import shutil
import subprocess

import pytest

NODE = shutil.which("node")


@pytest.mark.skipif(NODE is None, reason="node is not available")
def test_wizard_mandatory_field_mapping():
    proc = subprocess.run([NODE, "tests/wizard_steps_check.js"],
                          capture_output=True, text=True, timeout=60)
    print(proc.stdout)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert "ALL PASS" in proc.stdout
