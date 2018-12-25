from cx_Freeze import setup, Executable

setup(name = "POSTAGE" ,
      version = "0.1" ,
      description = "" ,
      executables = [Executable("mailMerge.py")])