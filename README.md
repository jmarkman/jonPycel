# jonPycel
My workplace modifications to [Greg Schultz's](https://github.com/gschultz49) Pycel

# Changes
* (largest) Divided the singular large .py file into a driver script and two package scripts
* Revamped all of the business-rule specific functions to conform with corporate desires as best as possible
* Added more business-rule specific functions
* Caught specific errors in try/except statements instead of leaving them open
* Added template switching ability
* Added templates for other SoVs and expanded template conversion libraries
* Cleaned up comment documentation
* Packaged in installer via Inno Setup for office distribution
* Tried to adhere to [PEP 8](https://www.python.org/dev/peps/pep-0008/)

# TODO
* Review Greg's export.py (renamed pycelexport.py for clarity) and implement it as part of Pycel's overall flow
* Implement try/catch for if the user runs a SoV through PyCel, and then runs that same SoV through it again without closing the results that are already open
* Review the template switching since it works on a try/catch and resuses the same block of code: this can 100% be made better but time will not allow for it for first release
* Implement Street 2 extraction - this will be a very large undertaking as it requires further rewriting of several functions for everything to work

# Worries
My biggest worry is that depending on the number and scale of feature requests, this is no longer going to be just a Python script. At some point, it's going to have to make the jump to full-fledged program so at the very least, the program makes sense from an outsider's perspective.