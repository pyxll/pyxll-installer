"""
Creates MSI installers for PyXLL projects using Wix.

Requires the Wix Toolset to be installed
http://wixtoolset.org/

Copy and edit the example config file for your own project.
Customizations can be made by copying and editing the templates.

Usage:
python make_installer -c config/example.yaml
"""
import subprocess
import itertools
import argparse
import logging
import fnmatch
import jinja2
import yaml
import uuid
import sys
import os
import re

_log = logging.getLogger()


def find_wix_toolset(config):
    """Find the Wix toolset commandline binaries"""
    bin_folder = config.get("wix", {}).get("bin", "")
    candle = "candle.exe"
    light = "light.exe"
    if bin_folder:
        if not os.path.exists(bin_folder):
            raise RuntimeError("Wix toolset bin folder '%s' does not exists. Check config." % bin_folder)
        candle = os.path.join(bin_folder, candle)
        light = os.path.join(bin_folder, light)

    pipe = subprocess.Popen([candle, "-?"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = pipe.communicate()
    if 0 != pipe.wait():
        _log.error(stderr)
        raise RuntimeError("Error running candle.exe. Check config and ensure Wix toolset is installed.")

    pipe = subprocess.Popen([light, "-?"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = pipe.communicate()
    if 0 != pipe.wait():
        _log.error(stderr)
        raise RuntimeError("Error running light.exe. Check config and ensure Wix toolset is installed.")

    return {
        "candle": candle,
        "light": light
    }


def get_template_args(config):
    """Gets the arguments for rendering the tempate wxs file"""
    kwargs = dict(config["installer"])
    kwargs["content_directories"] = content_directories = {}
    kwargs["content_components"] = content_components = {}
    used_ids = {}

    def make_id(name):
        """Create a unique id for a given name for use in a wxs file"""
        wix_id = re.sub("[^a-z0-9]", "_", name.lower())
        while "__" in wix_id:
            wix_id = wix_id.replace("__", "_")

        if len(wix_id) > 72:
            wix_id = wix_id[:72]

        number = 1
        base_id = wix_id
        while wix_id in used_ids and name != used_ids[wix_id]:
            wix_id = base_id[:71 - len(str(number))] + "_" + str(number)
            number += 1

        used_ids[wix_id] = name
        return wix_id

    def get_component(directory):
        """Gets or creates a component for a directory"""
        component = content_components.get(directory["id"])
        if component is not None:
            return component

        component = content_components[directory["id"]] = {
            "id": make_id(directory["id"] + "|Component"),
            "guid": str(uuid.uuid4()),
            "directory": directory,
            "files": []
        }

        return component

    def get_directory(path):
        """Gets or creates a directory from a path"""
        path = os.path.normpath(path)
        if not path.startswith("."):
            path = os.path.join(".", path)

        parent = content_directories
        curr_path = ""
        for component in path.split(os.path.sep):
            curr_path = os.path.normpath(os.path.join(curr_path, component)) if curr_path else component
            directory_id = make_id(curr_path) if curr_path != "." else "ADDINFOLDER"
            directory = parent.get(component)
            if directory is None:
                directory = {
                    "id": directory_id,
                    "remove_folder_id": make_id(directory_id + "|rmdir"),
                    "remove_file_id": make_id(directory_id + "|rm"),
                    "name": component if component != "." else "|PRODUCT_NAME|",
                    "children": {}
                }
                parent[component] = directory

            get_component(directory)  # ensure a component is created for each directory
            parent = directory["children"]

        return directory

    # get the files to add
    content_root = config["installer"]["content"]["root"]
    excludes = config["installer"]["content"].get("exclude", [])
    for root, dirs, files in os.walk(content_root):
        paths = [os.path.join(root, f) for f in files]
        for exclude in excludes:
            excluded = fnmatch.filter(paths, exclude)
            if excluded:
                paths = [p for p in paths if p not in set(excluded)]

        # Create the directory and component dicts
        directory = get_directory(os.path.relpath(root, content_root))
        component = get_component(directory)

        # Add the files to the component
        for path in paths:
            target = os.path.relpath(path, content_root)
            component["files"].append({
                "id": make_id(target),
                "name": os.path.basename(target),
                "source": path
            })

    # check the xll files exist
    xll_32 = config["installer"].get("xll_32")
    xll_64 = config["installer"].get("xll_64")
    if not xll_32 and not xll_64:
        raise RuntimeError("xll_32 and/or xll_64 must be specified")

    if xll_32 and not os.path.exists(os.path.join(content_root, xll_32)):
        raise RuntimeError("xll_32 file '%s' not found" % xll_32)

    if xll_64 and not os.path.exists(os.path.join(content_root, xll_64)):
        raise RuntimeError("xll_64 file '%s' not found" % xll_64)

    kwargs["xll_32"] = xll_32.replace("/", os.path.sep) if xll_32 else None
    kwargs["xll_64"] = xll_64.replace("/", os.path.sep) if xll_64 else None

    return kwargs


def build_installer(config, build_dir, wix_tools):
    """Build the installer using Wix"""
    if not os.path.exists(build_dir):
        os.makedirs(build_dir)

    # Generate the input files from the templates
    def render_template(input, params):
        if not os.path.exists(input):
            raise RuntimeError("Template file %s not found" % input)

        filename = os.path.basename(input)
        output = os.path.join(build_dir, filename)

        with open(input) as fh:
            template = jinja2.Template(fh.read())

        with open(output, "w") as fh:
            fh.write(template.render(**params))

        return output

    # Render the main product template
    template_args = get_template_args(config)
    input_file = render_template(config["installer"]["product_template"], template_args)

    # and any localizations
    localizations = []
    for loc in config["installer"].get("localizations", []):
        localizations.append(render_template(loc, template_args))

    # and any custom messages
    if ("custom_messages" in config["installer"]):
        localizations.append(render_template(config["installer"]["custom_messages"], template_args))

    # Wix extensions used to build the installer
    extensions = ["WixUIExtension", "WixNetFxExtension", "WixUtilExtension"]
    ext_args = list(itertools.chain(*(["-ext", ext] for ext in extensions)))

    # Use candle to compile the wxs file
    output_file = os.path.join(build_dir, os.path.splitext(os.path.basename(input_file))[0] + ".wixobj")
    subprocess.check_call([wix_tools["candle"], input_file, "-o", output_file] + ext_args)

    # And light to build the MSI
    product_msi = os.path.join(build_dir, config["installer"]["msi_filename"])
    loc_args = list(itertools.chain(*[["-loc", x] for x in localizations]))
    subprocess.check_call([wix_tools["light"], output_file, "-o", product_msi] + ext_args + loc_args)

    _log.info("Success! Finished Building installer '%s'" % product_msi)


def main():
    logging.basicConfig(level=logging.INFO)
    parser = argparse.ArgumentParser()
    parser.add_argument("-c", "--config", help="Path to config file", required=True)
    args = parser.parse_args()

    config = args.config
    if not os.path.exists(config):
        raise RuntimeError("Config file %s not found" % config)

    config_name = os.path.splitext(os.path.basename(config))[0]
    with open(config) as fh:
        config = yaml.safe_load(fh)

    # Get the Wix command line tools
    wix_tools = find_wix_toolset(config)

    # Build the installer
    build_dir = os.path.join("_build", config_name)
    build_installer(config, build_dir, wix_tools)


if __name__ == "__main__":
    sys.exit(main())
