"""DEV launcher for the Sims 3 CC Health Scanner shared core."""
import os

os.environ["SIMS3_CC_SCANNER_EDITION"] = "DEV"

# The scanner currently executes on import.  Keeping the implementation in one
# module prevents DEV and PUBLIC health/classification logic from drifting apart.
import Sims3_CC_Health_Scanner_CORE  # noqa: F401,E402
