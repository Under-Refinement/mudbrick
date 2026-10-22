"""Run ExtendScript (.jsx) in an Adobe app on macOS, via AppleScript.

Every tool that needs an Adobe app goes through run_jsx(), so the details of
talking to each app live in one place.
"""
import subprocess

# AppleScript bundle ids of the apps mudbrick can drive.
APPS = {
    "indesign": "com.adobe.InDesign",
}

_RUN_INDESIGN = """
on run argv
    set jsx to POSIX file (item 1 of argv)
    set jsxArgs to rest of argv
    with timeout of %d seconds
        tell application id "%s"
            return do script jsx language javascript with arguments jsxArgs
        end tell
    end timeout
end run
"""


class AdobeError(RuntimeError):
    pass


def run_jsx(app, script, args=(), timeout=3600):
    """Run the ExtendScript file `script` in `app` ("indesign"), passing `args`
    (available as `arguments` inside the script). Returns what the script's last
    expression evaluates to, as text."""
    if app not in APPS:
        raise AdobeError("mudbrick does not know how to drive %r yet (known: %s)"
                         % (app, ", ".join(APPS)))
    applescript = _RUN_INDESIGN % (timeout, APPS[app])
    proc = subprocess.run(["osascript", "-", str(script)] + [str(a) for a in args],
                          input=applescript, capture_output=True, text=True)
    if proc.returncode != 0:
        raise AdobeError(proc.stderr.strip() or "osascript failed")
    return proc.stdout.rstrip("\n")
