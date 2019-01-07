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
