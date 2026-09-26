"""Server configuration for integration tests.

!! Never use this configuration in production because it
opens the server to the world and provide access to JupyterLab
JavaScript objects through the global window variable.
"""
import os

from jupyterlab.galata import configure_jupyter_server

configure_jupyter_server(c)

# `or`, not a get() default: an exported but empty JUPYTER_TEST_PORT must fall
# back to 8888, as playwright.config.js does
c.ServerApp.port = int(os.environ.get("JUPYTER_TEST_PORT") or "8888")

# Uncomment to set server log level to debug level
# c.ServerApp.log_level = "DEBUG"
