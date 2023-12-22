from py2exe import freeze

freeze(
    console=[{"script": "createCheckInsDatasheet.py", "icon_resources": [],
              "dest_base": "Create Data Input File"}],
    windows=[],
    data_files=[
        [('inputs'), ['inputs/EnrollmentDetailRpt.xlsx']]
    ],
    zipfile='resources.zip',
    options={"excludes": ['tkinter', 'tkinter.constants'],
             "bundle_files": 0, "verbose": 4},
    version_info={}
)
