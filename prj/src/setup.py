from cx_Freeze import setup, Executable

setup(
    name="mailMerge",
    # options={"build_exe": {"packages": ["idna", "requests"]}},
    version="0.1",
    description="",
    executables=[
        Executable("Step_1_Filter_addresses.py"),
        Executable("Step_2_Make_labels.py"),
        Executable("Step_3A_Send_SG_emails.py"),
        Executable("Step_3B_Send_DHL_emails.py"),
    ],
)
