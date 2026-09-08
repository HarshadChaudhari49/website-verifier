# -*- coding: utf-8 -*-
"""
Shrink the ChatGPT Chrome profiles to just what keeps you signed in.

    python clean_profiles.py            # every profile
    python clean_profiles.py acct2      # one profile
    python clean_profiles.py --dry-run  # show what would go, delete nothing

WHY THIS EXISTS
---------------
Each saved account is a full Chrome user-data directory. The part that
holds the session is about 1 MB; the rest is cache, and it grows back
every single run -- 4.3 GB of it on one profile, mostly Chrome's
on-device AI model. Re-signing in by hand is the thing this project
most wants to avoid, so the cleaner works from a KEEP list, not a
delete list: anything not named below goes, and the login survives.

WHAT MUST BE KEPT, AND WHY
--------------------------
  Local State              On Windows this holds os_crypt.encrypted_key
                           -- the key the cookie VALUES are encrypted
                           with. Delete it and Cookies still exists but
                           cannot be decrypted, so you are logged out
                           with no obvious reason why. This is the
                           least obvious file in the whole profile and
                           the easiest one to lose.
  Default/Network/         Cookies + Cookies-journal: the session.
  Default/Local Storage/   chatgpt.com keeps auth state here too.
  Default/IndexedDB/       and here.
  Default/Preferences      profile settings Chrome expects to find.
  Default/Secure Preferences
  Default/Accounts/        the signed-in Google account.
  Default/Sync Data/

Everything else -- Cache, Code Cache, GPUCache, Service Worker,
Extensions, History, the AI models, shader caches -- is regenerable.
Chrome rebuilds what it needs on the next launch.

SAFETY
------
Refuses to run while any Chrome is open on these profiles: deleting
from a live profile corrupts it. Verify afterwards with

    python website_verifier3.py --chatgpt-profile <name>

which prints the signed-in account.
"""

import os
import shutil
import sys


PROFILES_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "chatgpt_profiles",
)

# Relative to one profile's root. Matched case-insensitively, and a
# name here protects that whole subtree.
KEEP = {
    "local state",
    "default",                       # the folder itself; pruned below
}

# Relative to <profile>/Default. Same rules.
KEEP_IN_DEFAULT = {
    "network",                       # Cookies live here
    "local storage",
    "indexeddb",
    "preferences",
    "secure preferences",
    "accounts",
    "sync data",
}


def _size(path):
    if os.path.isfile(path):
        try:
            return os.path.getsize(path)
        except OSError:
            return 0
    total = 0
    for root, _dirs, files in os.walk(path):
        for name in files:
            try:
                total += os.path.getsize(os.path.join(root, name))
            except OSError:
                pass
    return total


def _remove(path, dry_run):
    if dry_run:
        return True
    try:
        if os.path.isfile(path):
            os.remove(path)
        else:
            shutil.rmtree(path)
        return True
    except Exception as exc:
        print(f"      could not remove {os.path.basename(path)}: "
              f"{type(exc).__name__}")
        return False


def chrome_is_running():
    """True if any chrome.exe has one of these profiles open."""
    try:
        import subprocess
        out = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "@(Get-CimInstance Win32_Process -Filter \"Name='chrome.exe'\" |"
             " Where-Object { $_.CommandLine -like '*chatgpt_profiles*' }).Count"],
            capture_output=True, text=True, timeout=30,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return int((out.stdout or "0").strip() or 0) > 0
    except Exception:
        # Cannot tell -- assume the worst rather than corrupt a profile.
        return True


def clean_profile(profile_dir, dry_run=False):
    """Strip one profile to the keep list. Returns bytes freed."""
    name = os.path.basename(profile_dir)
    before = _size(profile_dir)
    freed = 0

    cookies = os.path.join(profile_dir, "Default", "Network", "Cookies")
    local_state = os.path.join(profile_dir, "Local State")
    if not os.path.isfile(cookies):
        print(f"  {name}: no Cookies file -- not signed in, skipping.")
        return 0
    if not os.path.isfile(local_state):
        print(f"  {name}: no Local State -- refusing, the cookies would")
        print("      not be decryptable afterwards.")
        return 0

    for entry in os.listdir(profile_dir):
        if entry.lower() in KEEP:
            continue
        path = os.path.join(profile_dir, entry)
        size = _size(path)
        if _remove(path, dry_run):
            freed += size

    default_dir = os.path.join(profile_dir, "Default")
    if os.path.isdir(default_dir):
        for entry in os.listdir(default_dir):
            if entry.lower() in KEEP_IN_DEFAULT:
                continue
            path = os.path.join(default_dir, entry)
            size = _size(path)
            if _remove(path, dry_run):
                freed += size

    after = before - freed
    verb = "would free" if dry_run else "freed"
    print(f"  {name:<10} {before/1048576:8.1f} MB -> {after/1048576:6.1f} MB"
          f"   {verb} {freed/1048576:.1f} MB")
    return freed


def main():
    args = [a for a in sys.argv[1:] if a != "--dry-run"]
    dry_run = "--dry-run" in sys.argv[1:]

    if not os.path.isdir(PROFILES_DIR):
        print(f"No profiles folder at {PROFILES_DIR}")
        return 1

    wanted = args or sorted(
        d for d in os.listdir(PROFILES_DIR)
        if os.path.isdir(os.path.join(PROFILES_DIR, d))
    )

    print("=" * 62)
    print("CHATGPT PROFILE CLEANER  [SYSTEM 3]"
          + ("   (DRY RUN)" if dry_run else ""))
    print("Keeps the session, deletes the cache.")
    print("=" * 62)

    if not dry_run and chrome_is_running():
        print("Chrome is open on one of these profiles -- close every")
        print("Chrome window first. Cleaning a live profile corrupts it.")
        return 1

    total = 0
    for name in wanted:
        path = os.path.join(PROFILES_DIR, name)
        if not os.path.isdir(path):
            print(f"  {name}: no such profile")
            continue
        total += clean_profile(path, dry_run)

    print("-" * 62)
    print(f"{'Would free' if dry_run else 'Freed'}: {total/1048576:.1f} MB")
    if not dry_run:
        print("Sessions are untouched. Verify with:")
        print("    python website_verifier3.py --chatgpt-profile <name>")
    return 0


if __name__ == "__main__":
    sys.exit(main())
