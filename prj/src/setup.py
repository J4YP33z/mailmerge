from cx_Freeze import setup, Executable

setup(
    name="mailMerge",
    version="0.1",
    description="",
    executables=[
        Executable("Step_1_Filter_addresses.py"),
        Executable("Step_2_Make_labels.py"),
    ],
)
