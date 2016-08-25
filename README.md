This project adds D bindings for libxlsxwriter (http://libxlsxwriter.github.io/index.html).

# Usage
This project is a `sourceLibrary` dub target type, so you should simply be able to add a dependency to your dub.json/dub.sdl file.  Note that you will need to install libxlsxwriter independently (see instructions here: http://libxlsxwriter.github.io/getting_started.html).

# Examples
A number of the examples included with libxlsxwriter have been translated to D and can be found in `examples/`.  These are minimal translations and strongly resemble the original C code.  You can build them like so:
```
$ dub build xlsxwriter:tutorial3
$ dub build xlsxwriter:chart
..etc..
```
