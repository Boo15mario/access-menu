import os
import zipfile

from SCons.Script import (
    AlwaysBuild,
    Alias,
    Builder,
    Clean,
    Default,
    DefaultEnvironment,
    Dir,
)

ROOT_DIR = Dir("#").abspath
ADDON_DIR = os.path.join(ROOT_DIR, "addon")
OUTPUT_NAME = "accessMenu.nvda-addon"
OUTPUT_PATH = os.path.join(ROOT_DIR, OUTPUT_NAME)


def _build_addon(target, source, env):
    if not os.path.isdir(ADDON_DIR):
        raise Exception(f"addon directory not found: {ADDON_DIR}")

    output_path = str(target[0])
    if os.path.exists(output_path):
        os.remove(output_path)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for root, _, files in os.walk(ADDON_DIR):
            for filename in files:
                file_path = os.path.join(root, filename)
                arc_name = os.path.relpath(file_path, ADDON_DIR)
                archive.write(file_path, arc_name)

    return None


env = DefaultEnvironment()
env.Append(BUILDERS={"AddonZip": Builder(action=_build_addon)})

def _collect_addon_sources(env):
    file_paths = []
    for root, _, files in os.walk(ADDON_DIR):
        for filename in files:
            file_paths.append(os.path.join(root, filename))
    file_paths.sort()
    return [env.File(path) for path in file_paths]


addon_sources = _collect_addon_sources(env)
target = env.AddonZip(OUTPUT_PATH, addon_sources)
Alias("build", target)
AlwaysBuild("build")
Default(target)
Clean(target, OUTPUT_PATH)
