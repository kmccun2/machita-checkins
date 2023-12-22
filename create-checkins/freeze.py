from py2exe import freeze

freeze(
    console=[{"script": "createCheckIns.py", "icon_resources": [],
              "dest_base": "Create Check-In Cards"}],
    windows=[],
    data_files=[
        [('inputs'), ['inputs/Check-In Cards (Data File).xlsx']],
        [('templates'), ['templates/Check-In Cards (template).xlsx']]
    ],
    zipfile='resources.zip',
    options={"excludes": ['tkinter', 'tkinter.constants'],
             "bundle_files": 0, "verbose": 4},
    version_info={}
)
