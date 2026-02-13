
# PyXLL Installer

This is an example project that builds an MSI installer for PyXLL based add-ins.

**Note:** PyXLL is licensed *per-user*. It is your responsibility to ensure that all users using your PyXLL
based add-in have a valid license. If you need to distribute your add-in outside of your organization where
you cannot control usage, speak to [sales@pyxll.com](mailto:sales@pyxll.com) about a redistributable
license.

The Wix Toolset is required to build MSI installers and may be downloaded from
[http://wixtoolset.org/](http://wixtoolset.org/).

WiX v6 has been tested and found to work correctly. If you are using a different version of WiX, you may need
to make modifications for it to work.

Using the WiX Toolset may require a maintenance agreement. You should check the WiX website for details to 
ensure you are compliant.

Use this project as a template. It is only an example intended to help get you started. You are expected to fork
this project and adapt the script, templates, and configuration to suit your own build environment and installer
requirements.

This project is based on the [Excel-DNA WiX Installer](https://github.com/Excel-DNA/WiXInstaller)
by [Govert van Drimmelen](https://github.com/govert).

## Introduction

There are various ways to distribute a PyXLL add-in, typically:

1. Sharing Python code on a network drive.
2. Bundling everything into a single zip file and scripting the installation.
3. Building an installer.

Option 1. is the simplest option as it allows your Python code to be deployed centrally, ensuring that all users
see the same code at all times.

Option 2. is the most commonly used option for large PyXLL deployments. It is less complicated than building
an MSI, can include the Python runtime as well as all dependencies, and automatic updates are possible 
using a PyXLL *startup script*.

Option 3. may be useful when you need to deploy to PCs that might not have fast or reliable access to your network,
and so accessing a shared drive is not feasible. The Python runtime can be bundled with PyXLL into a single
standalone installer.

If you need to distribute PyXLL outside of your organization, please make sure you have consulted with
[sales@pyxll.com](mailto:sales@pyxll.com) to ensure you have an appropriate license. PyXLL licenses are not
valid for redistributing PyXLL based add-ins, and you will need a separate license.

## Building a PyXLL Installer

Before starting to build your installer you need to decide the following:

1. Do you want to include Python in your installer?
2. Will you support 32 bit Excel, 64 bit Excel, or both?

If you do not include Python in your installer, your end users will have to have the correct version of Python
installed, including any dependencies not included in your installer.

If you intend to support both 32 bit Excel and 64 bit Excel (which is recommended) and you want to bundle
Python in the installer (which is also recommended) you will need both the 32 bit and 64 bit versions of
PyXLL and Python.

The basics steps you need to follow to build an installer are:

1. Copy your files and (optionally) your relocatable Python distributions to `content`.
2. Configure your PyXLL add-in correctly.
3. Create a new YAML config file under `config`.
4. Run `make_installer.py`.


### Content Directory Structure

**Note** If you are using `pyxll-bake` you should *not* include any config files in your installer. Use the baked
add-ins instead of the plain *pyxll.xll* files, which will include your config and any baked Python modules.

The content for your installer goes in the 'content' folder in this project. The suggested directory structure
is as follows:

```txt
content
|--- your_project
|     |    your_python_code  (folder)
|     |    shared_pyxll.cfg
|     |    ....
|     |
|     |--- x86
|     |    |    pyxll.xll
|     |    |    pyxll.cfg
|     |    |    logs  (folder)
|     |    |    PythonXX_x86  (folder)
|     |    |    ...
|     |
|     |--- x64
|     |    |    pyxll.xll
|     |    |    pyxll.cfg
|     |    |    logs  (folder)
|     |    |    PythonXX_x64  (folder)
|     |    |    ...
```

The pyxll.xll files in the x86 and x64 folders are the 32 bit and 64 bit versions respectively. Note that
you can rename these files however you like to make it easier to keep track of which is which, but if you
do that you also need to rename the pyxll.cfg files found next to them to match.

In the PythonXX_x86 and PythonXX_x64 folders you need a 32 bit and 64 bit relocatable Python runtime environment.
`conda-pack` is a convenient way to create a relocatable Python environment suitable for this. See
https://conda.github.io/conda-pack/ for instructions on how to use `conda-pack`. Note that PythonXX_x86 and
PythonXX_x64 are folders, and so if using `conda-pack` you will need to extract the tarball it creates.

The logs folders are included as that is the default place PyXLL will log to, and if the logs folder is
not included in the installer it may not be removed when uninstalled. If you intend to log elsewhere you
will not need the logs folders.

If you do not require the 32 bit or 64 bit version, you can simply exclude the one you don't want.


### PyXLL Configuration

**Note** If you are using `pyxll-bake` you should *not* include any config files in your installer. Use the baked
add-ins instead of the plain *pyxll.xll* files, which will include your config and any baked Python modules.

Each pyxll.xll (the x86 and x64 versions) should be configured to use the Python runtime relative to
the config file.

To make configuration easier it is suggested that you use an `external_config` as shown in the directory
structure above shared between the two different versions of the add-in. This is also relative to each pyxll.cfg
file.

*pyxll.cfg*:

```ini
[PYXLL]
external_config = ../shared_pyxll.cfg

[PYTHON]
executable = ./PythonXX_xZZ/pythonw.exe

[LOG]
path = %(YOUR_ADDIN_LOG_PATH:./logs)s
```

You main configuration will be in the *shared_pyxll.cfg* file. You can call this whatever you like, just be sure
to reference it correctly in the two *pyxll.cfg* files. This would include any settings from your normal pyxll.cfg file.

*shared_pyxll.cfg*:

```ini
[PYXLL]
modules =
    your_python_module

[PYTHON]
pythonpath =
    ./your_python_code

[LOG]
verbosity = %(YOUR_ADDIN_LOG_LEVEL:info)s
file = your-addin.%(date)s.log
```

In both these config files the `%(ENV_VAR:default)s` substitution syntax has been used. This allows your end users to
change certain properties, like log level and path, without having to modify any config.


## Installer Configuration

The `config` folder contains an example configuration file. Copy this and rename it to the name of your project
and edit it to your requirements.

At the very least you should set the following *installer* settings:

- **product_guid** Set to a new GUID for your project. GUIDs can be created by various different tools,
  including https://www.guidgenerator.com/.

- **product_version** Set this to your product version. You will need to increment this for each release you do.
  
- **msi_filename** This is the filename of the built msi.

- **localized_product_details** Set these to your own product name, description etc.

- **content.root** Change this to `content/your_product`, which is the root folder containing your files.
 
Additionally, you may wish to change `xll_32` and `xll_64` if you have renamed your `pyxll.xll` files. If you
only require one of these, you can remove whichever you don't need. 

The resources `product_icon`, `license_rtf`, `banner_bmp` and `dialog_bmp` should all be changed before finalizing
your installer to customize the appearance of the installer and to include your own license terms.

Depending on where you have installed the WiX Toolset, and what version you have, you may need to update
**wix.bin**.

## Building the Installer

The installer is built using the `make_installer.py` script.

Before you can run the script, install the dependencies listed in the `requirements.txt` file by running

```bat
conda install --file requirements.txt
```

Or if you are not using Anaconda use

```bat
pip install -r requirements.txt
```

Now you should be able to run the script. Note that you should be in the same folder as the script when you run it.

```code
cd pyxll-installer
python ./pyxll-installer -c ./config/my_project.yaml
```

If successful, your installer will now be in the `_build` folder!


## Advanced Customization and Localization

The WiX source files are in the `templates` folder. You can copy and edit these and update your YAML config
to refer to your own templates.

If you want to support multiple languages in your installed, create localization templates based on
`EnglishLoc.wxl` and add them to the list of localizations in your YAML config file.

Text used in the installer can be over-ridden with a `CustomMessages.wxl` file.
